// EmfVectorPdfWriter.cs
// Полноценный ВЕКТОРНЫЙ конвертер EMF -> PDF (полная переработка).
//
// Подход:
//   * EMF парсится через EnumEnhMetaFile (gdi32.dll) — GDI отдаёт нам сырые записи.
//   * Каждая запись транслируется в операторы PDF content stream.
//   * Текст — векторный, шрифт TrueType встраивается (WinAnsiEncoding + Differences для кириллицы).
//   * Линии/прямоугольники/полилинии/полигоны — чистый вектор.
//   * Изображения (EMR_STRETCHDIBITS) — встроенные XObject'ы (DeviceRGB).
//
// API совместим со сломанной версией:
//   public static bool Save(IList<byte[]> emfPages, string fileName, EmfPdfSettings settings = null)
//
// Зависимости: .NET Framework 4.0+, System.Drawing, gdi32.dll (Windows).
// Внешних NuGet-пакетов не требуется.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace Moxel
{
    // ======================================================================
    //  Публичный API
    // ======================================================================

    /// <summary>Настройки конвертации EMF -> PDF.</summary>
    public partial class EmfPdfSettings
    {
        /// <summary>Сжимать содержимое страниц (FlateDecode).</summary>
        public bool Compress { get; set; } = true;

        /// <summary>
        /// Полный путь к TTF-шрифту, который будет встроен в PDF.
        /// Если null — берётся системный Arial из %WINDIR%\Fonts. Шрифт должен
        /// поддерживать кириллицу.
        /// </summary>
        public string EmbeddedFontPath { get; set; }

        /// <summary>
        /// Полный путь к bold-варианту того же шрифта. 
        /// </summary>
        public string EmbeddedFontBoldPath { get; set; }

        /// <summary>
        /// Полный путь к italic-варианту того же шрифта. 
        /// </summary>
        public string EmbeddedFontItalicPath { get; set; }

        /// <summary>
        /// Полный путь к bolditalic-варианту того же шрифта. 
        /// </summary>
        public string EmbeddedFontBoldItalicPath { get; set; }

    }

    /// <summary>Фасад конвертера EMF -> PDF. Совместим со сломанной версией по сигнатуре.</summary>
    public static class EmfVectorPdfWriter
    {
        /// <summary>Сохраняет список EMF-страниц в один PDF-файл.</summary>
        public static bool Save(IList<byte[]> emfPages, string fileName, EmfPdfSettings settings = null)
        {
            if (emfPages == null || emfPages.Count == 0 || string.IsNullOrEmpty(fileName))
                return false;

            settings = settings ?? new EmfPdfSettings();

            try
            {
                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, 256 * 1024))
                {
                    var pdf = new PdfDocument(fs, settings);

                    string[] paths =
                    {
                        settings.EmbeddedFontPath        ?? ResolveSystemFont("arial.ttf"),
                        settings.EmbeddedFontBoldPath    ?? ResolveSystemFont("arialbd.ttf"),
                        settings.EmbeddedFontItalicPath  ?? ResolveSystemFont("ariali.ttf"),
                        settings.EmbeddedFontBoldItalicPath ?? ResolveSystemFont("arialbi.ttf"),
                    };

                    // fallback на regular/bold, если italic-файлов нет
                    paths[1] = paths[1] ?? paths[0];
                    paths[2] = paths[2] ?? paths[0];
                    paths[3] = paths[3] ?? paths[1];

                    var cache = new Dictionary<string, TtfFont>(StringComparer.OrdinalIgnoreCase);
                    var idCache = new Dictionary<TtfFont, int>();
                    var fonts = new TtfFont[4];
                    var fontIds = new int[4];
                    var synthItalic = new bool[4];
                    for (int i = 0; i < 4; i++)
                    {
                        if (!cache.TryGetValue(paths[i], out var f)) { f = TtfFont.Load(paths[i]); cache[paths[i]] = f; }
                        fonts[i] = f;
                        if (!idCache.TryGetValue(f, out int id)) { id = pdf.AddFont(f); idCache[f] = id; }
                        fontIds[i] = id;
                    }
                    // если italic-слот получил файл regular/bold — наклоняем синтетически
                    synthItalic[2] = ReferenceEquals(fonts[2], fonts[0]);
                    synthItalic[3] = ReferenceEquals(fonts[3], fonts[1]);

                    // Обрабатываем каждую EMF-страницу
                    foreach (var emfBytes in emfPages)
                    {
                        if (emfBytes == null || emfBytes.Length == 0) continue;
                        var page = EmfPageProcessor.Process(emfBytes, pdf, fonts, fontIds, synthItalic, settings);
                        if (page != null)
                            pdf.AddPage(page);
                    }

                    pdf.Finish();
                }
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("EmfVectorPdfWriter: " + ex);
                return false;
            }
        }

        private static string ResolveSystemFont(params string[] candidates)
        {
            string fontsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
            foreach (var c in candidates)
            {
                string p = Path.Combine(fontsDir, c);
                if (File.Exists(p)) return p;
            }
            return null;
        }
    }

    // ======================================================================
    //  Низкоуровневый писатель PDF
    // ======================================================================

    internal sealed class PdfDocument
    {
        private readonly Stream _fs;
        private readonly EmfPdfSettings _settings;
        private readonly Dictionary<int, long> _offsets = new Dictionary<int, long>();
        private readonly List<KeyValuePair<int, TtfFont>> _fonts = new List<KeyValuePair<int, TtfFont>>();
        private int _nextId = 1;

        public int CatalogId { get; }
        public int PagesId { get; }
        public List<int> PageIds { get; } = new List<int>();

        public PdfDocument(Stream fs, EmfPdfSettings settings)
        {
            _fs = fs;
            _settings = settings;
            WriteAscii("%PDF-1.4\n");
            WriteBytes(new byte[] { 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A });
            CatalogId = AllocId();
            PagesId = AllocId();
        }

        public int AllocId() => _nextId++;

        public void WriteAscii(string s) { var b = Encoding.ASCII.GetBytes(s); _fs.Write(b, 0, b.Length); }
        public void WriteBytes(byte[] b) { _fs.Write(b, 0, b.Length); }

        public void StartObject(int id) { _offsets[id] = _fs.Position; WriteAscii(id + " 0 obj\n"); }
        public void EndObject() { WriteAscii("\nendobj\n"); }

        public int AddFont(TtfFont font)
        {
            int id = AllocId();
            _fonts.Add(new KeyValuePair<int, TtfFont>(id, font));
            return id;
        }
        
        private readonly Dictionary<ulong, int> _imgCache = new Dictionary<ulong, int>();

        public int AddImage(byte[] imageData, int width, int height, bool isJpeg)
        {
            ulong hash = 1469598103934665603UL;                    // FNV-1a
            foreach (byte b in imageData) { hash ^= b; hash *= 1099511628211UL; }
            hash ^= (ulong)(uint)(width * 397 ^ height);
            if (_imgCache.TryGetValue(hash, out int cached)) 
                return cached;

            int imgId = AllocId();
            StartObject(imgId);
            WriteAscii("<< /Type /XObject /Subtype /Image /Width " + width + " /Height " + height);
            WriteAscii(" /ColorSpace /DeviceRGB /BitsPerComponent 8");
            WriteAscii(" /Filter /" + (isJpeg ? "DCTDecode" : "FlateDecode"));
            WriteAscii(" /Length " + imageData.Length + " >>\nstream\n");
            WriteBytes(imageData);
            WriteAscii("\nendstream");
            EndObject();
            _imgCache[hash] = imgId;
            return imgId;
        }

        public void AddPage(PdfPage page)
        {
            int pageId = AllocId();
            int contId = AllocId();
            PageIds.Add(pageId);

            // Сборка Resources
            string xobj = "";
            if (page.ImageRefs.Count > 0)
            {
                var sb = new StringBuilder("/XObject <<");
                for (int i = 0; i < page.ImageRefs.Count; i++)
                    sb.Append(" /Im" + (i + 1) + " " + page.ImageRefs[i] + " 0 R");
                sb.Append(" >>");
                xobj = sb.ToString();
            }

            // Page
            StartObject(pageId);
            WriteAscii("<< /Type /Page /Parent " + PagesId + " 0 R");
            WriteAscii(" /MediaBox [0 0 " + F(page.MediaWidth) + " " + F(page.MediaHeight) + "]");
            WriteAscii(" /Resources << /Font << /F1 " + page.FontIds[0] + " 0 R /F2 " + page.FontIds[1] + " 0 R /F3 " + page.FontIds[2] + " 0 R /F4 " + page.FontIds[3] + " 0 R >> " + xobj + " >>");
            WriteAscii(" /Contents " + contId + " 0 R >>");
            EndObject();
 
            byte[] raw = Encoding.GetEncoding("iso-8859-1").GetBytes(page.Content.ToString());
            byte[] content = _settings.Compress ? PdfZlib.Deflate(raw) : raw;

            StartObject(contId);
            WriteAscii("<< /Length " + content.Length);
            if (_settings.Compress) 
                WriteAscii(" /Filter /FlateDecode");
            WriteAscii(" >>\nstream\n");
            WriteBytes(content);
            WriteAscii("\nendstream");
            EndObject();
        }

        private void WriteType0Font(int fontId, TtfFont font)
        {
            if (font.FontProgram == null || font.FontProgram.Length == 0)
            {
                StartObject(fontId);
                WriteAscii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
                EndObject();
                return;
            }

            int fileId = AllocId(), descId = AllocId(), cidId = AllocId(), tuId = AllocId();

            byte[] prog = font.FontProgram;
            try { prog = TtfSubset.Build(font.FontProgram, font.UsedGlyphs.Keys); }
            catch (Exception ex) { Debug.WriteLine("TtfSubset: " + ex.Message); }
            byte[] comp = PdfZlib.Deflate(prog);   // ваш zlib-вариант с заголовком и adler32

            string baseFont = SubsetTag(font) + "+" + font.BaseFont;   // напр. KFDHBA+ArialMT

            StartObject(fileId);
            WriteAscii("<< /Length " + comp.Length + " /Filter /FlateDecode /Length1 " + prog.Length + " >>\nstream\n");
            WriteBytes(comp);
            WriteAscii("\nendstream");
            EndObject();


            //StartObject(fileId);
            //WriteAscii("<< /Length " + font.FontProgram.Length + " /Length1 " + font.FontProgram.Length + " >>\nstream\n");
            //WriteBytes(font.FontProgram);
            //WriteAscii("\nendstream");
            //EndObject();

            StartObject(descId);   // метрики уже нормированы к 1000/em (прошлый патч)
            WriteAscii("<< /Type /FontDescriptor /FontName /" + font.BaseFont);
            WriteAscii(" /Flags " + font.Flags);
            WriteAscii(" /FontBBox [" + font.BBox[0] + " " + font.BBox[1] + " " + font.BBox[2] + " " + font.BBox[3] + "]");
            WriteAscii(" /ItalicAngle " + font.ItalicAngle.ToString("0.#", CultureInfo.InvariantCulture));
            WriteAscii(" /Ascent " + font.Ascent + " /Descent " + font.Descent);
            WriteAscii(" /CapHeight " + font.CapHeight + " /StemV " + font.StemV);
            WriteAscii(" /FontFile2 " + fileId + " 0 R >>");
            EndObject();

            var w = new StringBuilder();
            foreach (var kv in font.UsedGlyphs)
            {
                int gw = kv.Key < font.GlyphWidths.Length ? font.GlyphWidths[kv.Key] : 1000;
                w.Append(kv.Key).Append(" [").Append(gw).Append("] ");
            }

            StartObject(cidId);
            WriteAscii("<< /Type /Font /Subtype /CIDFontType2 /BaseFont /" + font.BaseFont);
            WriteAscii(" /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >>");
            WriteAscii(" /FontDescriptor " + descId + " 0 R /DW 1000 /W [ " + w + "] /CIDToGIDMap /Identity >>");
            EndObject();

            byte[] tu = BuildToUnicode(font);
            StartObject(tuId);
            WriteAscii("<< /Length " + tu.Length + " >>\nstream\n");
            WriteBytes(tu);
            WriteAscii("\nendstream");
            EndObject();

            StartObject(fontId);
            WriteAscii("<< /Type /Font /Subtype /Type0 /BaseFont /" + font.BaseFont);
            WriteAscii(" /Encoding /Identity-H /DescendantFonts [" + cidId + " 0 R] /ToUnicode " + tuId + " 0 R >>");
            EndObject();
        }

        private static string SubsetTag(TtfFont f)
        {
            int h = 17;
            foreach (var g in f.UsedGlyphs.Keys) h = h * 31 + g;
            var sb = new StringBuilder(6);
            for (int i = 0; i < 6; i++) { sb.Append((char)('A' + (Math.Abs(h) % 26))); h /= 26; }
            return sb.ToString();
        }

        private static byte[] BuildToUnicode(TtfFont font)
        {
            var sb = new StringBuilder();
            sb.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n");
            sb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
            sb.Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n");
            sb.Append("1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");
            var items = new List<KeyValuePair<ushort, ushort>>(font.UsedGlyphs);
            for (int i = 0; i < items.Count; i += 100)
            {
                int n = Math.Min(100, items.Count - i);
                sb.Append(n).Append(" beginbfchar\n");
                for (int j = 0; j < n; j++)
                    sb.Append("<").Append(items[i + j].Key.ToString("X4"))
                      .Append("> <").Append(((int)items[i + j].Value).ToString("X4")).Append(">\n");
                sb.Append("endbfchar\n");
            }
            sb.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend");
            return Encoding.ASCII.GetBytes(sb.ToString());
        }

        public void Finish()
        {
            foreach (var kv in _fonts) 
                WriteType0Font(kv.Key, kv.Value);
            // Catalog
            StartObject(CatalogId);
            WriteAscii("<< /Type /Catalog /Pages " + PagesId + " 0 R >>");
            EndObject();

            // Pages
            StartObject(PagesId);
            WriteAscii("<< /Type /Pages /Count " + PageIds.Count + " /Kids [");
            foreach (var id in PageIds) WriteAscii(id + " 0 R ");
            WriteAscii("] >>");
            EndObject();

            // Xref
            long xrefPos = _fs.Position;
            WriteAscii("xref\n0 " + _nextId + "\n");
            WriteAscii("0000000000 65535 f \n");
            for (int i = 1; i < _nextId; i++)
            {
                if (_offsets.TryGetValue(i, out long off))
                    WriteAscii(off.ToString("D10") + " 00000 n \n");
                else
                    WriteAscii("0000000000 65535 f \n");
            }
            WriteAscii("trailer\n<< /Size " + _nextId + " /Root " + CatalogId + " 0 R >>\n");
            WriteAscii("startxref\n" + xrefPos + "\n%%EOF");
        }

        private static string F(double d) => d.ToString("0.####", CultureInfo.InvariantCulture);
    }

    /// <summary>Одна страница PDF: размер + content stream + ссылки на изображения.</summary>
    internal sealed class PdfPage
    {
        public double MediaWidth;
        public double MediaHeight;
        public StringBuilder Content = new StringBuilder();
        public List<int> ImageRefs = new List<int>();
        internal int[] FontIds;
    }

    /// <summary>RGB-изображение (24-bit, без альфы) для встраивания в PDF как XObject.</summary>
    internal sealed class RgbImage
    {
        public byte[] Data;
        public int Width;
        public int Height;
    }

    // ======================================================================
    //  Zlib (Deflate) — для /FlateDecode
    // ======================================================================

    internal static class PdfZlib
    {
        // .NET DeflateStream выдаёт «сырой» zlib-deflate (без zlib-заголовка и adler32).
        // PDF для /FlateDecode ждёт именно такой формат (RFC 1950 zlib без 2-байт
        // заголовка и 4-байт adler32 — это и есть raw deflate).
        public static byte[] Deflate(byte[] data)
        {
            using (var ms = new MemoryStream())
            {
                ms.WriteByte(0x78); // CMF: deflate, окно 32K
ms.WriteByte(0x9C); // FLG: контрольная сумма сходимости
                using (var ds = new DeflateStream(ms, CompressionMode.Compress, true))
                {
                    ds.Write(data, 0, data.Length);
                    ds.Flush();
                }
                return ms.ToArray();
            }
        }
    }

    // ======================================================================
    //  TTF-парсер и встраивание шрифта
    // ======================================================================

    internal sealed class TtfFont
    {
        public string BaseFont;            // Имя для /BaseFont (например, "ArialMT")
        public byte[] FontProgram;         // Содержимое .ttf файла
        public int Flags = 32;             // Regular; 32|64 = bold; 32|96 = bold italic; 32|64? - см. PDF spec
        public int Ascent;
        public int Descent;
        public int CapHeight;
        public int StemV = 80;
        public double ItalicAngle = 0;
        public int[] BBox = new int[4];     // [xMin yMin xMax yMax] в design units
        public int WinAscent;   // usWinAscent, per-1000 после нормализации

        public Dictionary<ushort, ushort> Cmap = new Dictionary<ushort, ushort>();
        public int[] GlyphWidths = new int[0];                                    // per-1000, индекс = GID
        internal short UnderlinePos;
        internal short UnderlineThickness;
        public readonly SortedDictionary<ushort, ushort> UsedGlyphs
            = new SortedDictionary<ushort, ushort>();                             // GID -> Unicode


        public static TtfFont Load(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return CreateFallback();
            try
            {
                byte[] data = File.ReadAllBytes(path);
                return new TtfParser(data).BuildFont();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("TtfFont.Load: " + ex);
                return CreateFallback();
            }
        }

        private static TtfFont CreateFallback()
        {
            // Если TTF недоступен — вернуть вырожденный шрифт. PDF откроется, но текст
            // может рендериться стандартным Helvetica без кириллицы.
            return new TtfFont
            {
                BaseFont = "Helvetica",
                FontProgram = new byte[0],
                Ascent = 905,
                Descent = -212,
                CapHeight = 728,
                StemV = 80,
                BBox = new int[] { -665, -325, 2029, 1037 },
            };
        }
    }

    internal sealed class TtfParser
    {
        private readonly byte[] _data;
        private readonly Dictionary<string, uint> _tableOff = new Dictionary<string, uint>();
        private readonly Dictionary<string, uint> _tableLen = new Dictionary<string, uint>();

        public TtfParser(byte[] data)
        {
            _data = data;
            ParseDirectory();
        }

        private void ParseDirectory()
        {
            if (_data.Length < 12) throw new InvalidDataException("TTF too small");
            ushort numTables = ReadU16(4);
            for (int i = 0; i < numTables; i++)
            {
                uint rec = 12 + (uint)(i * 16);
                string tag = ReadTag(rec);
                _tableOff[tag] = ReadU32(rec + 8);
                _tableLen[tag] = ReadU32(rec + 12);
            }
        }

        private ushort ReadU16(uint o) => (ushort)((_data[o] << 8) | _data[o + 1]);
        private short ReadI16(uint o) => (short)ReadU16(o);
        private uint ReadU32(uint o) => ((uint)_data[o] << 24) | ((uint)_data[o + 1] << 16) | ((uint)_data[o + 2] << 8) | _data[o + 3];
        private string ReadTag(uint o) => Encoding.ASCII.GetString(_data, (int)o, 4);

        public TtfFont BuildFont()
        {
            var f = new TtfFont { FontProgram = _data };

            // head: unitsPerEm, bbox
            if (_tableOff.TryGetValue("head", out var hOff))
            {
                f.BBox[0] = ReadI16(hOff + 36);
                f.BBox[1] = ReadI16(hOff + 38);
                f.BBox[2] = ReadI16(hOff + 40);
                f.BBox[3] = ReadI16(hOff + 42);
            }

            if (_tableOff.TryGetValue("post", out var postOff))
            {
                f.UnderlinePos = ReadI16(postOff + 8);        // обычно отрицательное (ниже baseline)
                f.UnderlineThickness = ReadI16(postOff + 10);
            }

            uint unitsPerEm = 1000;
            if (_tableOff.TryGetValue("head", out var h2))
                unitsPerEm = ReadU16(h2 + 18);

            // hhea: ascent/descent, numberOfHMetrics
            uint numHMetrics = 0;
            if (_tableOff.TryGetValue("hhea", out var hOff2))
            {
                f.Ascent = ReadI16(hOff2 + 4);
                f.Descent = ReadI16(hOff2 + 6);
                numHMetrics = ReadU16(hOff2 + 34);
            }

            // OS/2: sTypoAscender, sTypoDescender, fsSelection (bold/italic), sCapHeight
            if (_tableOff.TryGetValue("OS/2", out var os2Off))
            {
                int os2Len = (int)_tableLen["OS/2"];
                if (os2Len >= 78)
                {
                    // sTypoAscender @68, sTypoDescender @70, sTypoLineGap @72, sFamilyClass @8, sCapHeight @88
                    short sTypoAsc = ReadI16(os2Off + 68);
                    short sTypoDesc = ReadI16(os2Off + 70);
                    if (sTypoAsc != 0 || sTypoDesc != 0) { f.Ascent = sTypoAsc; f.Descent = sTypoDesc; }
                    if (os2Len >= 90)
                    {
                        short capH = ReadI16(os2Off + 88);
                        f.CapHeight = capH != 0 ? capH : (int)(f.Ascent * 0.7);
                    }
                    else
                        f.CapHeight = (int)(f.Ascent * 0.7);
                    // fsSelection @62: bit0=italic, bit5=bold
                    ushort fsSel = ReadU16(os2Off + 62);
                    if ((fsSel & 0x20) != 0) f.Flags |= 1 << 18;              // ForceBold
                    if ((fsSel & 0x01) != 0) { f.Flags |= 64; f.ItalicAngle = -12; } // Italic
                    f.WinAscent = ReadU16(os2Off + 74);   // usWinAscent
                }
            }
            if (f.CapHeight == 0) f.CapHeight = (int)(f.Ascent * 0.7);
            if (f.StemV == 0) f.StemV = (f.Flags & 64) != 0 ? 120 : 80;

            // name:family — для BaseFont
            if (_tableOff.TryGetValue("name", out var nameOff))
            {
                string baseName = SanitizeFontName(ReadName(nameOff) ?? "EmbeddedFont");
                if ((f.Flags & 64) != 0) baseName += "Bold";
                if ((f.Flags & 4) != 0) baseName += "Italic";
                f.BaseFont = baseName;
            }
            else
            {
                f.BaseFont = "EmbeddedFont";
            }

            // cmap: char -> glyphId
            Dictionary<ushort, ushort> cmap = ParseCmap();

            f.Cmap = cmap;

            ushort numGlyphs = 0;
            if (_tableOff.TryGetValue("maxp", out var maxpOff))
                numGlyphs = ReadU16(maxpOff + 4);

            if (_tableOff.TryGetValue("hmtx", out var hmtxOff) && numHMetrics > 0 && numGlyphs > 0)
            {
                f.GlyphWidths = new int[numGlyphs];
                for (ushort gid = 0; gid < numGlyphs; gid++)
                {
                    ushort w = GetHmtxWidth(hmtxOff, numHMetrics, gid);
                    f.GlyphWidths[gid] = (int)Math.Round(1000.0 * w / unitsPerEm);
                }
            }

            double k = 1000.0 / unitsPerEm;
            f.Ascent = (int)Math.Round(f.Ascent * k);
            f.Descent = (int)Math.Round(f.Descent * k);
            f.CapHeight = (int)Math.Round(f.CapHeight * k);
            f.WinAscent = (int)Math.Round(f.WinAscent * k);
            f.UnderlinePos = (short)Math.Round(f.UnderlinePos * k);
            f.UnderlineThickness = (short)Math.Round(f.UnderlineThickness * k);
            if (f.WinAscent == 0) 
                f.WinAscent = f.Ascent;
            for (int i = 0; i < 4; i++) 
                f.BBox[i] = (int)Math.Round(f.BBox[i] * k);

            return f;
        }

        private string SanitizeFontName(string s)
        {
            var sb = new StringBuilder();
            foreach (char ch in s)
                if (char.IsLetterOrDigit(ch)) sb.Append(ch);
            return sb.Length > 0 ? sb.ToString() : "EmbeddedFont";
        }

        private string ReadName(uint nameOff)
        {
            ushort format = ReadU16(nameOff);
            ushort count = ReadU16(nameOff + 2);
            uint storageOff = ReadU16(nameOff + 4);
            for (int i = 0; i < count; i++)
            {
                uint rec = nameOff + 6 + (uint)(i * 12);
                ushort platform = ReadU16(rec);
                ushort encoding = ReadU16(rec + 2);
                ushort language = ReadU16(rec + 4);
                ushort nameID = ReadU16(rec + 6);
                ushort length = ReadU16(rec + 8);
                ushort offset = ReadU16(rec + 10);
                if (nameID == 1) // Family
                {
                    uint full = nameOff + storageOff + offset;
                    // Защита от битых шрифтов: строка должна лежать внутри буфера.
                    if (full + length > (uint)_data.Length) return null;
                    // platform 0 (Unicode) и platform 3/encoding 1 (Windows BMP) — UTF-16BE
                    if (platform == 0 || (platform == 3 && encoding == 1))
                        return ReadUtf16Be(full, length / 2);
                    else
                        return Encoding.GetEncoding("iso-8859-1").GetString(_data, (int)full, length);
                }
            }
            return null;
        }

        private string ReadUtf16Be(uint off, int count)
        {
            if (off + count * 2 > _data.Length) return "";
            var sb = new StringBuilder(count);
            for (int i = 0; i < count; i++)
                sb.Append((char)ReadU16(off + (uint)(i * 2)));
            return sb.ToString();
        }

        private Dictionary<ushort, ushort> ParseCmap()
        {
            var result = new Dictionary<ushort, ushort>();
            if (!_tableOff.TryGetValue("cmap", out var cmapOff)) return result;

            ushort numSub = ReadU16(cmapOff + 2);
            uint chosenSub = 0;
            for (int i = 0; i < numSub; i++)
            {
                uint rec = cmapOff + 4 + (uint)(i * 8);
                ushort platform = ReadU16(rec);
                ushort encoding = ReadU16(rec + 2);
                uint subOff = cmapOff + ReadU32(rec + 4);
                // Windows BMP Unicode (3,1) или Microsoft Unicode UCS-4 (3,10)
                if (platform == 3 && (encoding == 1 || encoding == 10))
                {
                    chosenSub = subOff;
                    if (encoding == 1) break;
                }
            }
            if (chosenSub == 0) return result;

            ushort format = ReadU16(chosenSub);
            if (format == 4) ParseCmapFormat4(chosenSub, result);
            else if (format == 6) ParseCmapFormat6(chosenSub, result);
            else if (format == 12) ParseCmapFormat12(chosenSub, result);
            return result;
        }

        private void ParseCmapFormat4(uint off, Dictionary<ushort, ushort> result)
        {
            ushort segX2 = ReadU16(off + 6);
            int segCount = segX2 / 2;
            uint endOff = off + 14;
            uint startOff = endOff + (uint)segX2 + 2;
            uint deltaOff = startOff + (uint)segX2;
            uint rangeOffOff = deltaOff + (uint)segX2;

            for (int i = 0; i < segCount; i++)
            {
                ushort end = ReadU16(endOff + (uint)(i * 2));
                ushort start = ReadU16(startOff + (uint)(i * 2));
                short delta = ReadI16(deltaOff + (uint)(i * 2));
                ushort rangeOff = ReadU16(rangeOffOff + (uint)(i * 2));
                if (start == 0xFFFF && end == 0xFFFF) continue;

                for (int ch = start; ch <= end; ch++)
                {
                    ushort gid;
                    if (rangeOff == 0)
                    {
                        gid = (ushort)((ch + delta) & 0xFFFF);
                    }
                    else
                    {
                        uint p = rangeOffOff + (uint)(i * 2) + rangeOff + (uint)((ch - start) * 2);
                        gid = ReadU16(p);
                        if (gid != 0) gid = (ushort)((gid + delta) & 0xFFFF);
                    }
                    if (gid != 0 && !result.ContainsKey((ushort)ch))
                        result[(ushort)ch] = gid;
                }
            }
        }

        private void ParseCmapFormat6(uint off, Dictionary<ushort, ushort> result)
        {
            ushort first = ReadU16(off + 6);
            ushort count = ReadU16(off + 8);
            uint p = off + 10;
            for (int ch = 0; ch < count; ch++)
            {
                ushort gid = ReadU16(p + (uint)(ch * 2));
                if (gid != 0) result[(ushort)(first + ch)] = gid;
            }
        }

        private void ParseCmapFormat12(uint off, Dictionary<ushort, ushort> result)
        {
            uint numGroups = ReadU32(off + 12);
            uint p = off + 16;
            for (uint g = 0; g < numGroups; g++)
            {
                uint startChar = ReadU32(p + g * 12);
                uint endChar = ReadU32(p + g * 12 + 4);
                uint startGid = ReadU32(p + g * 12 + 8);
                for (uint ch = startChar; ch <= endChar && ch <= 0xFFFF; ch++)
                {
                    ushort gid = (ushort)(startGid + (ch - startChar));
                    if (!result.ContainsKey((ushort)ch))
                        result[(ushort)ch] = gid;
                }
            }
        }

        private ushort GetHmtxWidth(uint hmtxOff, uint numHMetrics, ushort gid)
        {
            if (gid < numHMetrics)
                return ReadU16(hmtxOff + (uint)(gid * 4));
            if (numHMetrics > 0)
                return ReadU16(hmtxOff + (uint)((numHMetrics - 1) * 4));
            return 0;
        }

        // CP1251 byte → Unicode. Используется при извлечении ширины глифов:
        // для кода 0..255 находим Unicode, дальше через cmap — glyph ID.
        private static int Cp1251ToUnicode(int c)
        {
            if (c < 0x80) return c;
            switch (c)
            {
                case 0xA1: return 0x040E; // Ў
                case 0xA2: return 0x045E; // ў
                case 0xA3: return 0x0402; // Ђ
                case 0xA4: return 0x0452; // ђ
                case 0xA5: return 0x0405; // Ѕ
                case 0xA6: return 0x0409; // Љ (в CP1251 тут буква, не Latin-1 ¦)
                case 0xA7: return 0x040A; // Њ
                case 0xA8: return 0x0401; // Ё
                case 0xA9: return 0x00A9; // ©
                case 0xAA: return 0x0404; // Є
                case 0xAB: return 0x00AB; // «
                case 0xAC: return 0x00AC; // ¬
                case 0xAD: return 0x00AD; // soft hyphen
                case 0xAE: return 0x00AE; // ®
                case 0xAF: return 0x0407; // Ї
                case 0xB0: return 0x00B0; // °
                case 0xB1: return 0x00B1; // ±
                case 0xB2: return 0x0406; // І
                case 0xB3: return 0x0456; // і
                case 0xB4: return 0x0491; // ґ
                case 0xB5: return 0x0455; // ѕ
                case 0xB6: return 0x00B6; // ¶
                case 0xB7: return 0x00B7; // ·
                case 0xB8: return 0x0451; // ё
                case 0xB9: return 0x2116; // №
                case 0xBA: return 0x0454; // є
                case 0xBB: return 0x00BB; // »
                case 0xBC: return 0x0458; // ј
                case 0xBD: return 0x0405; // Ѕ (в CP1251 тут Ѕ, не ½)
                case 0xBE: return 0x0455; // ѕ
                case 0xBF: return 0x0457; // ї
                default:
                    if (c >= 0xC0 && c <= 0xFF) return 0x0410 + (c - 0xC0); // А..я
                    return c;
            }
        }
    }

    internal static class TtfSubset
    {
        public static byte[] Build(byte[] src, ICollection<ushort> usedGlyphs)
        {
            ushort numTables = RU16(src, 4);
            var off = new Dictionary<string, int>();
            var len = new Dictionary<string, int>();
            for (int i = 0; i < numTables; i++)
            {
                int rec = 12 + i * 16;
                string tag = Encoding.ASCII.GetString(src, rec, 4);
                off[tag] = (int)RU32(src, rec + 8);
                len[tag] = (int)RU32(src, rec + 12);
            }
            if (!off.ContainsKey("glyf") || !off.ContainsKey("loca")) return src;

            int headOff = off["head"];
            ushort numGlyphs = RU16(src, off["maxp"] + 4);
            bool longLoca = RU16(src, headOff + 50) != 0;

            var loca = new int[numGlyphs + 1];
            int lo = off["loca"];
            for (int i = 0; i <= numGlyphs; i++)
                loca[i] = longLoca ? (int)RU32(src, lo + i * 4) : RU16(src, lo + i * 2) * 2;

            // Замыкание по компонентам составных глифов
            int glyfOff = off["glyf"];
            var keep = new HashSet<ushort> { 0 };
            var stack = new Stack<ushort>(usedGlyphs);
            while (stack.Count > 0)
            {
                ushort gid = stack.Pop();
                if (gid >= numGlyphs || !keep.Add(gid)) continue;
                if (loca[gid + 1] - loca[gid] < 10) continue;
                int go = glyfOff + loca[gid];
                if ((short)RU16(src, go) >= 0) continue;   // простой глиф
                int p = go + 10;
                while (true)
                {
                    ushort flags = RU16(src, p);
                    stack.Push(RU16(src, p + 2));
                    p += 4 + ((flags & 0x0001) != 0 ? 4 : 2);      // args
                    if ((flags & 0x0008) != 0) p += 2;             // scale
                    else if ((flags & 0x0040) != 0) p += 4;        // x&y scale
                    else if ((flags & 0x0080) != 0) p += 8;        // 2x2
                    if ((flags & 0x0020) == 0) break;              // more components
                }
            }

            // Новый glyf + loca (нумерация сохранена)
            var glyf = new MemoryStream();
            var newLoca = new int[numGlyphs + 1];
            for (int gid = 0; gid < numGlyphs; gid++)
            {
                newLoca[gid] = (int)glyf.Length;
                int l = loca[gid + 1] - loca[gid];
                if (l > 0 && keep.Contains((ushort)gid))
                {
                    glyf.Write(src, glyfOff + loca[gid], l);
                    while (glyf.Length % 4 != 0) glyf.WriteByte(0);
                }
            }
            newLoca[numGlyphs] = (int)glyf.Length;

            var locaB = new byte[(numGlyphs + 1) * 4];
            for (int i = 0; i <= numGlyphs; i++) WU32(locaB, i * 4, (uint)newLoca[i]);

            var headB = Slice(src, headOff, len["head"]);
            WU32(headB, 8, 0);                 // checkSumAdjustment
            headB[50] = 0; headB[51] = 1;      // indexToLocFormat = long

            string[] want = { "cvt ", "fpgm", "glyf", "head", "hhea", "hmtx", "loca", "maxp", "prep" };
            var tables = new List<KeyValuePair<string, byte[]>>();
            foreach (var tag in want)
            {
                if (tag == "glyf") tables.Add(new KeyValuePair<string, byte[]>(tag, glyf.ToArray()));
                else if (tag == "loca") tables.Add(new KeyValuePair<string, byte[]>(tag, locaB));
                else if (tag == "head") tables.Add(new KeyValuePair<string, byte[]>(tag, headB));
                else if (off.ContainsKey(tag))
                    tables.Add(new KeyValuePair<string, byte[]>(tag, Slice(src, off[tag], len[tag])));
            }

            int n = tables.Count;
            int entrySel = (int)Math.Floor(Math.Log(n, 2));
            int searchRange = (1 << entrySel) * 16;
            using (var outMs = new MemoryStream())
            {
                var h = new byte[12];
                WU32(h, 0, 0x00010000); WU16(h, 4, (ushort)n);
                WU16(h, 6, (ushort)searchRange); WU16(h, 8, (ushort)entrySel);
                WU16(h, 10, (ushort)(n * 16 - searchRange));
                outMs.Write(h, 0, 12);

                var dir = new byte[n * 16];
                var data = new MemoryStream();
                int dataOff = 12 + n * 16;
                for (int i = 0; i < n; i++)
                {
                    byte[] t = tables[i].Value;
                    Encoding.ASCII.GetBytes(tables[i].Key).CopyTo(dir, i * 16);
                    WU32(dir, i * 16 + 4, Checksum(t));
                    WU32(dir, i * 16 + 8, (uint)(dataOff + data.Length));
                    WU32(dir, i * 16 + 12, (uint)t.Length);
                    data.Write(t, 0, t.Length);
                    while (data.Length % 4 != 0) data.WriteByte(0);
                }
                outMs.Write(dir, 0, dir.Length);
                data.Position = 0; data.CopyTo(outMs);
                return outMs.ToArray();
            }
        }

        private static ushort RU16(byte[] d, int o) => (ushort)((d[o] << 8) | d[o + 1]);
        private static uint RU32(byte[] d, int o) => ((uint)d[o] << 24) | ((uint)d[o + 1] << 16) | ((uint)d[o + 2] << 8) | d[o + 3];
        private static void WU16(byte[] d, int o, ushort v) { d[o] = (byte)(v >> 8); d[o + 1] = (byte)v; }
        private static void WU32(byte[] d, int o, uint v) { d[o] = (byte)(v >> 24); d[o + 1] = (byte)(v >> 16); d[o + 2] = (byte)(v >> 8); d[o + 3] = (byte)v; }
        private static byte[] Slice(byte[] d, int o, int l) { var r = new byte[l]; Buffer.BlockCopy(d, o, r, 0, l); return r; }
        private static uint Checksum(byte[] t)
        {
            uint s = 0;
            for (int i = 0; i < t.Length; i += 4)
                s += ((uint)t[i] << 24)
                   | (i + 1 < t.Length ? (uint)t[i + 1] << 16 : 0)
                   | (i + 2 < t.Length ? (uint)t[i + 2] << 8 : 0)
                   | (i + 3 < t.Length ? t[i + 3] : 0u);
            return s;
        }
    }

    // ======================================================================
    //  P/Invoke: GDI32 для работы с EMF
    // ======================================================================

    internal static class Gdi32
    {
        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern IntPtr SetEnhMetaFileBits(uint bufSize, byte[] buffer);

        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern bool DeleteEnhMetaFile(IntPtr hemf);

        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern uint GetEnhMetaFileHeader(IntPtr hemf, uint cbBuffer, byte[] buffer);

        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern bool EnumEnhMetaFile(IntPtr hdc, IntPtr hemf, EnumProc lpProc, IntPtr lpData, IntPtr lpRect);

        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern bool PlayEnhMetaFileRecord(IntPtr hdc, IntPtr lpht, IntPtr lprc, int nHandles);

        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern uint GetObjectType(IntPtr hgdiobj);

        [DllImport("gdi32.dll", EntryPoint = "GetObjectW", SetLastError = true)]
        public static extern int GetObject(IntPtr hgdiobj, int cbBuffer, byte[] lpvObject);

        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        // Callback отдаёт IntPtr на сырую запись, мы парсим сами через Marshal.
        public delegate bool EnumProc(IntPtr hdc, IntPtr ht, IntPtr pRecord, int nHandles, IntPtr data);

        public const uint OBJ_PEN = 1;
        public const uint OBJ_BRUSH = 2;
        public const uint OBJ_FONT = 6;
        public const int LOGPIXELSX = 88;
        public const int LOGPIXELSY = 90;
    }

    // ENHMETAHEADER — частично. Нам нужны rclFrame и rclBounds, signature, и размер.
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    internal struct EnhmHeader
    {
        public int iType;            // EMR_HEADER = 1
        public uint nSize;
        public int rclBoundsLeft;
        public int rclBoundsTop;
        public int rclBoundsRight;
        public int rclBoundsBottom;
        public int rclFrameLeft;
        public int rclFrameTop;
        public int rclFrameRight;
        public int rclFrameBottom;
        public uint dSignature;       // ENHMETA_SIGNATURE = 0x464D4520
        public uint nVersion;
        public uint nBytes;
        public uint nRecords;
        public ushort nHandles;
        public ushort sReserved;
        public uint nDescription;
        public uint offDescription;
        public uint nPalEntries;
        public int szlDeviceCx;
        public int szlDeviceCy;
        public int szlMillimetersCx;
        public int szlMillimetersCy;
    }

    // ======================================================================
    //  Парсер одной EMF-страницы -> PdfPage
    // ======================================================================
    internal sealed class EmfPageProcessor
    {
        private sealed class DcState
        {
            public uint PenColor, BrushColor, TextColor;
            public float PenWidth;
            public bool PenNull, BrushNull, FontBold, FontItalic;
            public double FontH;
            public int PosX, PosY;
        }

        // Типы записей EMF (MS-EMF). Взято только нужное.
        private const int
            EMR_HEADER = 1,
            EMR_POLYBEZIER = 2,
            EMR_POLYGON = 3,
            EMR_POLYLINE = 4,
            EMR_POLYBEZIERTO = 5,
            EMR_POLYLINETO = 6,
            EMR_POLYPOLYLINE = 7,
            EMR_POLYPOLYGON = 8,
            EMR_SETWINDOWEXTEX = 9,
            EMR_SETWINDOWORGEX = 10,
            EMR_SETVIEWPORTEXTEX = 11,
            EMR_SETVIEWPORTORGEX = 12,
            EMR_SETBRUSHORGEX = 13,
            EMR_EOF = 14,
            EMR_SETPIXELV = 15,
            EMR_SETMAPPERFLAGS = 16,
            EMR_SETMAPMODE = 17,
            EMR_SETBKMODE = 18,
            EMR_SETPOLYFILLMODE = 19,
            EMR_SETROP2 = 20,
            EMR_SETSTRETCHBLTMODE = 21,
            EMR_SETTEXTALIGN = 22,
            EMR_SETCOLORADJUSTMENT = 23,
            EMR_SETTEXTCOLOR = 24,
            EMR_SETBKCOLOR = 25,
            EMR_OFFSETCLIPRGN = 26,
            EMR_MOVETOEX = 27,
            EMR_SETMETARGN = 28,
            EMR_EXCLUDECLIPRECT = 29,
            EMR_INTERSECTCLIPRECT = 30,
            EMR_SCALEVIEWPORTEXTEX = 31,
            EMR_SCALEWINDOWEXTEX = 32,
            EMR_SAVEDC = 33,
            EMR_RESTOREDC = 34,
            EMR_SETWORLDTRANSFORM = 35,
            EMR_MODIFYWORLDTRANSFORM = 36,
            EMR_SELECTOBJECT = 37,
            EMR_CREATEPEN = 38,
            EMR_CREATEBRUSHINDIRECT = 39,
            EMR_DELETEOBJECT = 40,
            EMR_ANGLEARC = 41,
            EMR_ELLIPSE = 42,
            EMR_RECTANGLE = 43,
            EMR_ROUNDRECT = 44,
            EMR_ARC = 45,
            EMR_CHORD = 46,
            EMR_PIE = 47,
            EMR_SELECTPALETTE = 48,
            EMR_CREATEPALETTE = 49,
            EMR_SETPALETTEENTRIES = 50,
            EMR_RESIZEPALETTE = 51,
            EMR_REALIZEPALETTE = 52,
            EMR_EXTFLOODFILL = 53,
            EMR_LINETO = 54,
            EMR_ARCTO = 55,
            EMR_POLYDRAW = 56,
            EMR_SETARCDIRECTION = 57,
            EMR_SETMITERLIMIT = 58,
            EMR_BEGINPATH = 59,
            EMR_ENDPATH = 60,
            EMR_CLOSEFIGURE = 61,
            EMR_FILLPATH = 62,
            EMR_STROKEANDFILLPATH = 63,
            EMR_STROKEPATH = 64,
            EMR_FLATTENPATH = 65,
            EMR_WIDENPATH = 66,
            EMR_SELECTCLIPPATH = 67,
            EMR_ABORTPATH = 68,
            EMR_COMMENT = 70,
            EMR_COMMENT_EMFPLUS = 0x2B464D45,
            EMR_COMMENT_EMFSPOOL = 0x00000000,
            EMR_COMMENT_PUBLIC = 0x43494447,
            EMR_FILLRGN = 71,
            EMR_FRAMERGN = 72,
            EMR_INVERTRGN = 73,
            EMR_PAINTRGN = 74,
            EMR_EXTSELECTCLIPRGN = 75,
            EMR_BITBLT = 76,
            EMR_STRETCHBLT = 77,
            EMR_MASKBLT = 78,
            EMR_PLGBLT = 79,
            EMR_SETDIBITSTODEVICE = 80,
            EMR_STRETCHDIBITS = 81,
            EMR_EXTCREATEFONTINDIRECTW = 82,
            EMR_EXTTEXTOUTA = 83,
            EMR_EXTTEXTOUTW = 84,
            EMR_POLYBEZIER16 = 85,
            EMR_POLYGON16 = 86,
            EMR_POLYLINE16 = 87,
            EMR_POLYBEZIERTO16 = 88,
            EMR_POLYLINETO16 = 89,
            EMR_POLYPOLYLINE16 = 90,
            EMR_POLYPOLYGON16 = 91,
            EMR_POLYDRAW16 = 92,
            EMR_CREATEMONOBRUSH = 93,
            EMR_CREATEDIBPATTERNBRUSHPT = 94,
            EMR_EXTCREATEPEN = 95,
            EMR_POLYTEXTOUTA = 96,
            EMR_POLYTEXTOUTW = 97,
            EMR_SETICMMODE = 98,
            EMR_CREATECOLORSPACE = 99,
            EMR_SETCOLORSPACE = 100,
            EMR_DELETECOLORSPACE = 101,
            EMR_GLSRECORD = 102,
            EMR_GLSBOUNDEDRECORD = 103,
            EMR_PIXELFORMAT = 104,
            EMR_DRAWESCAPE = 105,
            EMR_EXTESCAPE = 106,
            EMR_STARTDOC = 107,
            EMR_SMALLTEXTOUT = 108,
            EMR_FORCEUFIMAPPING = 109,
            EMR_NAMEDESCAPE = 110,
            EMR_COLORCORRECTPALETTE = 111,
            EMR_SETICMPROFILEA = 112,
            EMR_SETICMPROFILEW = 113,
            EMR_ALPHABLEND = 114,
            EMR_ALPHADIBBLEND = 115,
            EMR_TRANSPARENTBLT = 116,
            EMR_TRANSPARENTDIB = 117,
            EMR_GRADIENTFILL = 118,
            EMR_SETLINKEDUFIS = 119,
            EMR_SETTEXTJUSTIFICATION = 120,
            PDF_SIGNATURE = 0x50444620;


        // Текущее состояние GDI
        private uint _textColor = 0x000000;
        private uint _bkColor = 0xFFFFFF;
        private uint _penColor = 0x000000;
        private float _penWidth = 0.5f;
        private uint _brushColor = 0xFFFFFF;
        private bool _brushIsNull = false;
        private bool _penIsNull = false;
        private double _fontSize = 12;     // в PDF points
        private bool _fontBold = false;
        private bool _fontItalic = false;
        private string _fontFace = "Arial";
        private bool _fontUnderline;
        private bool _fontStrike;
        private double _fontEmHeightInPx = 16;

        // Текущая позиция пера (для EMR_MOVETOEX/EMR_LINETO)
        private int _curPosX = 0;
        private int _curPosY = 0;
        private bool _hasMoveTo = false;

        // World transform: 2x3 matrix in PDF order (a b c d e f)
        // EMF units are pixels at the EMF DPI; PDF content stream is in points.
        // Мы выставляем cm( scale 0 0 -scale 0 pageH ) и затем применяем world transform.
        private double _mxA = 1, _mxB = 0, _mxC = 0, _mxD = 1, _mxE = 0, _myF = 0;
        private uint _textAlign;
        private readonly PdfDocument _pdf;
        private readonly EmfPdfSettings _settings;
        private readonly bool[] _synthItalic;
        private readonly TtfFont[] _fonts;
        private readonly int[] _fontIds;
        private readonly IntPtr _hEmf;
        private readonly PdfPage _page = new PdfPage();

        private readonly List<int> _imageIdsOnPage = new List<int>();

        private abstract class GdiObj { }
        private sealed class PenObj : GdiObj { public uint Color; public float Width; public bool IsNull; }
        private sealed class BrushObj : GdiObj { public uint Color; public bool IsNull; }
        private sealed class FontObj : GdiObj { public double HeightPx; public bool Bold, Italic, Underline, StrikeOut; public string Face; }

        private readonly Dictionary<uint, GdiObj> _objects = new Dictionary<uint, GdiObj>();

        private readonly Stack<DcState> _dcStack = new Stack<DcState>();

        private EmfPageProcessor(IntPtr hEmf, PdfDocument pdf, TtfFont[] fonts, int[] fontIds, bool[] synthItalic, EmfPdfSettings settings)
        {
            _hEmf = hEmf;
            _pdf = pdf;
            _settings = settings;

            _synthItalic = synthItalic;
            _fonts = fonts;
            _fontIds = fontIds;

            _page.FontIds = fontIds;
        }

        public static PdfPage Process(byte[] emfBytes, PdfDocument pdf, TtfFont[] fonts, int[] fontIds, bool[] synthItalic, EmfPdfSettings settings)
        {
            IntPtr hEmf = Gdi32.SetEnhMetaFileBits((uint)emfBytes.Length, emfBytes);
            if (hEmf == IntPtr.Zero) return null;
            try
            {
                var proc = new EmfPageProcessor(hEmf, pdf, fonts, fontIds, synthItalic, settings);
                return proc.Run();
            }
            finally
            {
                Gdi32.DeleteEnhMetaFile(hEmf);
            }
        }

        private double _scaleX, _scaleY;

        private PdfPage Run()
        {
            // Читаем заголовок
            byte[] hdrBuf = new byte[108];
            uint got = Gdi32.GetEnhMetaFileHeader(_hEmf, (uint)hdrBuf.Length, hdrBuf);
            if (got == 0) return null;
            EnhmHeader hdr;
            // Пиннинг буфера без unsafe-кода.
            GCHandle pin = GCHandle.Alloc(hdrBuf, GCHandleType.Pinned);
            try
            {
                hdr = (EnhmHeader)Marshal.PtrToStructure(pin.AddrOfPinnedObject(), typeof(EnhmHeader));
            }
            finally
            {
                pin.Free();
            }
            if (hdr.dSignature != 0x464D4520) return null;

            // Размер страницы: rclFrame — в 0.01 mm.
            double mmW = Math.Max(1, hdr.rclFrameRight - hdr.rclFrameLeft) / 100.0;
            double mmH = Math.Max(1, hdr.rclFrameBottom - hdr.rclFrameTop) / 100.0;
            _page.MediaWidth = mmW * 72.0 / 25.4;
            _page.MediaHeight = mmH * 72.0 / 25.4;

            // Фактическое разрешение устройства записи EMF
            double dpiX = (double)hdr.szlDeviceCx * 25.4 / Math.Max(1, hdr.szlMillimetersCx);
            double dpiY = (double)hdr.szlDeviceCy * 25.4 / Math.Max(1, hdr.szlMillimetersCy);

            // Settings.Dpi остаётся только как fallback на случай мусора в заголовке
            if (dpiX < 30 || dpiX > 4800) dpiX = Math.Max(50, _settings.Dpi);
            if (dpiY < 30 || dpiY > 4800) dpiY = dpiX;

            _scaleX = 72.0 / dpiX;
            _scaleY = 72.0 / dpiY;

            // Пролог content stream: q + cm
            _page.Content.Append("q\n");
            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                "{0:0.######} 0 0 {1:0.######} 0 0 cm\n", _scaleX, _scaleY);
            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                "1 0 0 -1 0 {0:0.####} cm\n", _page.MediaHeight / _scaleY);

            // Перечисляем записи
            var gh = GCHandle.Alloc(this);
            try
            {
                Gdi32.EnumEnhMetaFile(IntPtr.Zero, _hEmf, CallbackProc, GCHandle.ToIntPtr(gh), IntPtr.Zero);
            }
            finally
            {
                gh.Free();
            }

            // Эпилог
            _page.Content.Append("Q\n");
            return _page;
        }

        private bool CallbackProc(IntPtr hdc, IntPtr ht, IntPtr pRec, int nHandles, IntPtr data)
        {
            int type = Marshal.ReadInt32(pRec);
            uint size = (uint)Marshal.ReadInt32(pRec, 4);
            try
            {
                HandleRecord(hdc, ht, pRec, nHandles, type, size);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("EMF record " + type + ": " + ex.Message);
            }
            return true;
        }

        // -------------------- Диспетчер --------------------

        private void HandleRecord(IntPtr hdc, IntPtr ht, IntPtr pRec, int nHandles, int type, uint size)
        {
            switch (type)
            {
                case EMR_SETTEXTCOLOR: _textColor = ReadU32(pRec, 8); break;
                case EMR_SETBKCOLOR: _bkColor = ReadU32(pRec, 8); break;
                case EMR_SETBKMODE: break; // игнор — у нас прозрачный фон
                case EMR_SETTEXTALIGN: _textAlign = ReadU32(pRec, 8); break;
                case EMR_SETWORLDTRANSFORM: ReadWorldTransform(pRec, 8); ApplyWorld(); break;
                case EMR_MODIFYWORLDTRANSFORM: HandleModifyWorld(pRec); break;
                //case EMR_SELECTOBJECT: HandleSelectObject(ht, pRec, nHandles); break;
                case EMR_SELECTOBJECT: HandleSelectObject(pRec); break;
                // === Создающие записи ОБЯЗАТЕЛЬНО отдавать в PlayEnhMetaFileRecord ===
                // EnumEnhMetaFile НЕ создаёт объекты сам: ячейки ht[] остаются нулями,
                // пока create-запись не проиграна через PEMR. Из-за пропуска этих
                // записей SELECTOBJECT раньше всегда читал IntPtr.Zero.
                case EMR_CREATEPEN: HandleCreatePen(pRec); break;
                case EMR_CREATEBRUSHINDIRECT: HandleCreateBrush(pRec); break;
                case EMR_EXTCREATEFONTINDIRECTW: HandleCreateFont(pRec); break;
                case EMR_EXTCREATEPEN: HandleExtCreatePen(pRec); break;
                case EMR_CREATEMONOBRUSH:
                case EMR_CREATEDIBPATTERNBRUSHPT:
                    // PEMR вызовет CreatePenIndirect / CreateBrushIndirect /
                    // CreateFontIndirectW / ExtCreatePen и запишет хэндл в ht[ih].
                    // DC для этих записей не нужен, hdc == Zero допустим.
                    Gdi32.PlayEnhMetaFileRecord(hdc, ht, pRec, nHandles);
                    break;
                case EMR_DELETEOBJECT: _objects.Remove(ReadU32(pRec, 8)); break;
                case EMR_RECTANGLE: HandleRectangle(pRec); break;
                case EMR_POLYGON16: HandlePoly(pRec, true); break;
                case EMR_POLYGON: HandlePoly32(pRec, true, false); break;
                case EMR_POLYLINE: HandlePoly32(pRec, false, false); break;
                case EMR_POLYLINETO: HandlePoly32(pRec, false, true); break;

                case EMR_POLYLINE16: HandlePoly(pRec, false); break;
                case EMR_POLYLINETO16: HandlePoly16To(pRec); break;
                case EMR_POLYPOLYLINE16: HandlePolyPoly16(pRec, false); break;
                case EMR_POLYPOLYGON16: HandlePolyPoly16(pRec, true); break;
                case EMR_POLYBEZIER16: HandleBezier16(pRec); break;
                case EMR_BITBLT: HandleBitBlt(pRec); break;
                case EMR_MOVETOEX: HandleMoveTo(pRec); break;
                case EMR_LINETO: HandleLineTo(pRec); break;
                case EMR_EXTTEXTOUTW: HandleText(hdc, pRec, size); break;
                case EMR_STRETCHDIBITS: HandleStretchDibits(pRec, size); break;
                case EMR_BEGINPATH:
                case EMR_ENDPATH:
                case EMR_FILLPATH:
                case EMR_STROKEPATH:
                case EMR_STROKEANDFILLPATH:
                    // Упрощённо игнорируем — EMF из 1С их почти не использует
                    break;
                case EMR_SAVEDC:
                    _page.Content.Append("q\n");
                    _dcStack.Push(new DcState
                    {
                        PenColor = _penColor,
                        BrushColor = _brushColor,
                        TextColor = _textColor,
                        PenWidth = _penWidth,
                        PenNull = _penIsNull,
                        BrushNull = _brushIsNull,
                        FontBold = _fontBold,
                        FontItalic = _fontItalic,
                        FontH = _fontEmHeightInPx,
                        PosX = _curPosX,
                        PosY = _curPosY
                    });
                    break;

                case EMR_RESTOREDC:
                    {
                        // nSavedDC @8: обычно -1; -n означает "снять n уровней"
                        int n = Marshal.ReadInt32(pRec, 8);
                        int count = n < 0 ? -n : Math.Max(1, _dcStack.Count - n + 1);
                        for (int i = 0; i < count && _dcStack.Count > 0; i++)
                        {
                            _page.Content.Append("Q\n");
                            var s = _dcStack.Pop();
                            _penColor = s.PenColor; _brushColor = s.BrushColor; _textColor = s.TextColor;
                            _penWidth = s.PenWidth; _penIsNull = s.PenNull; _brushIsNull = s.BrushNull;
                            _fontBold = s.FontBold; _fontItalic = s.FontItalic; _fontEmHeightInPx = s.FontH;
                            _curPosX = s.PosX; _curPosY = s.PosY;
                        }
                        break;
                    }

                case EMR_INTERSECTCLIPRECT:
                    {
                        // RECTL @8: left, top, right, bottom
                        int l = Marshal.ReadInt32(pRec, 8);
                        int t = Marshal.ReadInt32(pRec, 12);
                        int r = Marshal.ReadInt32(pRec, 16);
                        int b = Marshal.ReadInt32(pRec, 20);
                        _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                            "{0} {1} {2} {3} re W n\n", l, t, r - l, b - t);
                        break;
                    }

                case EMR_SETWINDOWORGEX:
                    {
                        // POINTL @8: последующие логические координаты сдвинуты на -org
                        int ox = Marshal.ReadInt32(pRec, 8);
                        int oy = Marshal.ReadInt32(pRec, 12);
                        if (ox != 0 || oy != 0)
                            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                                "1 0 0 1 {0} {1} cm\n", -ox, -oy);
                        break;
                    }

                case EMR_SETVIEWPORTORGEX:
                    {
                        int ox = Marshal.ReadInt32(pRec, 8);
                        int oy = Marshal.ReadInt32(pRec, 12);
                        if (ox != 0 || oy != 0)
                            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                                "1 0 0 1 {0} {1} cm\n", ox, oy);
                        break;
                    }

                default:
                    Debug.WriteLine($"Не поддерживаемый тип записи EMF: {type}");
                    break;
            }
        }

        private void HandleCreatePen(IntPtr pRec)
        {
            // EMR_CREATEPEN: ihPen@8, LOGPEN { style@12, width.x@16, width.y@20, color@24 }
            uint ih = ReadU32(pRec, 8);
            uint style = ReadU32(pRec, 12);
            int width = Marshal.ReadInt32(pRec, 16);
            uint color = ReadU32(pRec, 24);
            _objects[ih] = new PenObj { Color = color, Width = Math.Max(1, width), IsNull = (style & 0xF) == 5 };
        }

        private void HandleExtCreatePen(IntPtr pRec)
        {
            // EMR_EXTCREATEPEN: ihPen@8, offBmi/cbBmi/offBits/cbBits@12..27,
            // EXTLOGPEN { penStyle@28, width@32, brushStyle@36, color@40 }
            uint ih = ReadU32(pRec, 8);
            uint style = ReadU32(pRec, 28);
            int width = Marshal.ReadInt32(pRec, 32);
            uint color = ReadU32(pRec, 40);
            _objects[ih] = new PenObj { Color = color, Width = Math.Max(1, width), IsNull = (style & 0xF) == 5 };
        }

        private void HandleCreateBrush(IntPtr pRec)
        {
            // EMR_CREATEBRUSHINDIRECT: ihBrush@8, LOGBRUSHEX { style@12, color@16, hatch@20 }
            uint ih = ReadU32(pRec, 8);
            uint style = ReadU32(pRec, 12);
            uint color = ReadU32(pRec, 16);
            _objects[ih] = new BrushObj { Color = color, IsNull = style == 1 }; // BS_NULL
        }

        private void HandleCreateFont(IntPtr pRec)
        {
            // EMR_EXTCREATEFONTINDIRECTW: ihFont@8, LOGFONTW@12:
            //   lfHeight@12, lfWidth@16, lfEscapement@20, lfOrientation@24,
            //   lfWeight@28, lfItalic@32, lfFaceName (32 WCHAR) @40
            uint ih = ReadU32(pRec, 8);
            int lfHeight = Marshal.ReadInt32(pRec, 12);
            int lfWeight = Marshal.ReadInt32(pRec, 28);
            byte lfItalic = Marshal.ReadByte(pRec, 32);
            byte lfUnderlined = Marshal.ReadByte(pRec, 33);
            byte lfStrikeOut = Marshal.ReadByte(pRec, 34);
            var sb = new StringBuilder();
            for (int i = 0; i < 32; i++)
            {
                char ch = (char)Marshal.ReadInt16(pRec, 40 + i * 2);
                if (ch == 0) break;
                sb.Append(ch);
            }
            _objects[ih] = new FontObj
            {
                HeightPx = Math.Abs(lfHeight),   // в EMF device units, НЕ переводим в points
                Bold = lfWeight >= 700,
                Italic = lfItalic != 0,
                Underline = lfUnderlined != 0,
                StrikeOut = lfStrikeOut != 0,
                Face = sb.ToString()
            };
        }

        // -------------------- Объекты --------------------
        private void HandleSelectObject(IntPtr pRec)
        {
            uint ih = ReadU32(pRec, 8);
            if ((ih & 0x80000000) != 0)  // stock object
            {
                switch (ih & 0x7FFFFFFF)
                {
                    case 0: _brushColor = 0xFFFFFF; _brushIsNull = false; break; // WHITE_BRUSH
                    case 1: _brushColor = 0xC0C0C0; _brushIsNull = false; break; // LTGRAY_BRUSH
                    case 2: _brushColor = 0x808080; _brushIsNull = false; break; // GRAY_BRUSH
                    case 3: _brushColor = 0x404040; _brushIsNull = false; break; // DKGRAY_BRUSH
                    case 4: _brushColor = 0x000000; _brushIsNull = false; break; // BLACK_BRUSH
                    case 5: _brushIsNull = true; break;                          // NULL_BRUSH
                    case 6: _penColor = 0xFFFFFF; _penWidth = 1; _penIsNull = false; break; // WHITE_PEN
                    case 7: _penColor = 0x000000; _penWidth = 1; _penIsNull = false; break; // BLACK_PEN
                    case 8: _penIsNull = true; break;                            // NULL_PEN
                }
                return;
            }
            if (!_objects.TryGetValue(ih, out var obj)) return;
            if (obj is PenObj p) { _penColor = p.Color; _penWidth = p.Width; _penIsNull = p.IsNull; }
            else if (obj is BrushObj br) { _brushColor = br.Color; _brushIsNull = br.IsNull; }
            else if (obj is FontObj fo)
            {
                _fontEmHeightInPx = fo.HeightPx > 0 ? fo.HeightPx : 16;
                _fontBold = fo.Bold; 
                _fontItalic = fo.Italic; 
                _fontFace = fo.Face;
                _fontUnderline = fo.Underline;
                _fontStrike = fo.StrikeOut;
            }
        }

        // -------------------- World transform --------------------

        private void ReadWorldTransform(IntPtr pRec, int off)
        {
            _mxA = ReadFloat(pRec, off);
            _mxB = ReadFloat(pRec, off + 4);
            _mxC = ReadFloat(pRec, off + 8);
            _mxD = ReadFloat(pRec, off + 12);
            _mxE = ReadFloat(pRec, off + 16);
            _myF = ReadFloat(pRec, off + 20);
        }

        private void HandleModifyWorld(IntPtr pRec)
        {
            // EMR_MODIFYWORLDTRANSFORM: iType, nSize, M11..M22, eX, eY, iMode
            // iMode @ 36: 1=MWT_IDENTITY, 2=MWT_LEFTMULTIPLY, 3=MWT_RIGHTMULTIPLY
            //uint mode = ReadU32(pRec, 36);
            uint mode = ReadU32(pRec, 32);
            var n11 = ReadFloat(pRec, 8);
            var n12 = ReadFloat(pRec, 12);
            var n21 = ReadFloat(pRec, 16);
            var n22 = ReadFloat(pRec, 20);
            var nX = ReadFloat(pRec, 24);
            var nY = ReadFloat(pRec, 28);

            if (mode == 1) // MWT_IDENTITY
            {
                _mxA = 1; _mxB = 0; _mxC = 0; _mxD = 1; _mxE = 0; _myF = 0;
            }
            else if (mode == 2) // LEFT
            {
                // New = New * Old
                var (a, b, c, d, e, f) = Mul(n11, n12, n21, n22, nX, nY, _mxA, _mxB, _mxC, _mxD, _mxE, _myF);
                _mxA = a; _mxB = b; _mxC = c; _mxD = d; _mxE = e; _myF = f;
            }
            else if (mode == 3) // RIGHT
            {
                var (a, b, c, d, e, f) = Mul(_mxA, _mxB, _mxC, _mxD, _mxE, _myF, n11, n12, n21, n22, nX, nY);
                _mxA = a; _mxB = b; _mxC = c; _mxD = d; _mxE = e; _myF = f;
            }
            ApplyWorld();
        }

        private static (double a, double b, double c, double d, double e, double f) Mul(
            double a1, double b1, double c1, double d1, double e1, double f1,
            double a2, double b2, double c2, double d2, double e2, double f2)
        {
            // (a1 b1 c1 d1 e1 f1) * (a2 b2 c2 d2 e2 f2) — обычное матричное умножение
            return (
                a1 * a2 + c1 * b2,
                b1 * a2 + d1 * b2,
                a1 * c2 + c1 * d2,
                b1 * c2 + d1 * d2,
                a1 * e2 + c1 * f2 + e1,
                b1 * e2 + d1 * f2 + f1
            );
        }

        private void ApplyWorld()
        {
            if (_dcStack.Count > 0)
            {
                Debug.WriteLine("SetWorldTransform внутри SaveDC — пропущен");
                return;
            }

            // Сбросить предыдущий мир через Q/q
            _page.Content.Append("Q q ");

            // Базовый масштаб заново
            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                "{0:0.######} 0 0 {1:0.######} 0 0 cm ", _scaleX, _scaleY);

            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                "1 0 0 -1 0 {0:0.####} cm ", _page.MediaHeight / _scaleY);

            // Затем world
            if (_mxA != 1 || _mxB != 0 || _mxC != 0 || _mxD != 1 || _mxE != 0 || _myF != 0)
            {
                _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                    "{0:0.#####} {1:0.#####} {2:0.#####} {3:0.#####} {4:0.#####} {5:0.#####} cm\n",
                    _mxA, _mxB, _mxC, _mxD, _mxE, _myF);
            }
            else
            {
                _page.Content.Append("\n");
            }
            // Восстановить текущие цвета после Q (pen=RG, brush=rg; textColor
            // устанавливается явно в HandleText)
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} RG\n", ColorToPdf(_penColor));
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} rg\n", ColorToPdf(_brushColor));
            //_page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0:0.##} w\n", _penWidth);
        }

        // -------------------- Геометрия --------------------

        private void HandleRectangle(IntPtr pRec)
        {
            int l = Marshal.ReadInt32(pRec, 8);
            int t = Marshal.ReadInt32(pRec, 12);
            int r = Marshal.ReadInt32(pRec, 16);
            int b = Marshal.ReadInt32(pRec, 20);
            double w = r - l;
            double h = b - t;
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} {2:0.###} {3:0.###} re\n", (double)l, (double)t, w, h);
            FillStroke(!_brushIsNull, !_penIsNull);
            //if (!_brushIsNull) Fill();
            //if (!_penIsNull) Stroke();
        }

        private void HandlePoly(IntPtr pRec, bool closed)
        {
            // EMR_POLYGON16/EMR_POLYLINE16:
            //   iType, nSize, rclBounds(16 bytes), cptl (4 bytes), aptl[POINTS (USHORT x, USHORT y)]
            //   точки по 4 байта (x:SHORT, y:SHORT)
            uint cptl = ReadU32(pRec, 24);
            int pointsOff = 28;
            var sb = new StringBuilder();
            for (uint i = 0; i < cptl; i++)
            {
                short x = Marshal.ReadInt16(pRec, pointsOff + (int)(i * 4));
                short y = Marshal.ReadInt16(pRec, pointsOff + (int)(i * 4) + 2);
                if (i == 0)
                    sb.AppendFormat(CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} m\n", (double)x, (double)y);
                else
                    sb.AppendFormat(CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} l\n", (double)x, (double)y);
            }
            _page.Content.Append(sb);
            if (closed) FillStroke(!_brushIsNull, !_penIsNull);
            else Stroke();
        }

        // 32-битные точки: POINTL по 8 байт, счётчик @24, точки @28
        private void HandlePoly32(IntPtr pRec, bool closed, bool fromCurrent)
        {
            uint cptl = ReadU32(pRec, 24);
            if (cptl == 0) return;
            if (fromCurrent)
                _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} {1} m\n", _curPosX, _curPosY);
            int lastX = _curPosX, lastY = _curPosY;
            for (uint i = 0; i < cptl; i++)
            {
                int x = Marshal.ReadInt32(pRec, 28 + (int)(i * 8));
                int y = Marshal.ReadInt32(pRec, 28 + (int)(i * 8) + 4);
                _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} {1} {2}\n",
                    x, y, (i == 0 && !fromCurrent) ? "m" : "l");
                lastX = x; lastY = y;
            }
            if (closed) { _page.Content.Append("h\n"); FillStroke(!_brushIsNull, !_penIsNull); }
            else Stroke();
            if (fromCurrent) { _curPosX = lastX; _curPosY = lastY; }
        }

        private void HandlePoly16To(IntPtr pRec)
        {
            uint cpts = ReadU32(pRec, 24);
            if (cpts == 0) return;
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} {1} m\n", _curPosX, _curPosY);
            for (uint i = 0; i < cpts; i++)
            {
                short x = Marshal.ReadInt16(pRec, 28 + (int)(i * 4));
                short y = Marshal.ReadInt16(pRec, 28 + (int)(i * 4) + 2);
                _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} {1} l\n", x, y);
                _curPosX = x; _curPosY = y;
            }
            Stroke();
        }

        // nPolys@24, cptsTotal@28, aPolyCounts[nPolys]@32, затем точки по 4 байта
        private void HandlePolyPoly16(IntPtr pRec, bool closed)
        {
            uint nPolys = ReadU32(pRec, 24);
            int ptsOff = 32 + (int)(nPolys * 4);
            int p = 0;
            for (uint k = 0; k < nPolys; k++)
            {
                uint cnt = ReadU32(pRec, 32 + (int)(k * 4));
                for (uint i = 0; i < cnt; i++, p++)
                {
                    short x = Marshal.ReadInt16(pRec, ptsOff + p * 4);
                    short y = Marshal.ReadInt16(pRec, ptsOff + p * 4 + 2);
                    _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} {1} {2}\n", x, y, i == 0 ? "m" : "l");
                }
                if (closed) _page.Content.Append("h\n");
            }
            if (closed) FillStroke(!_brushIsNull, !_penIsNull);
            else Stroke();
        }

        // PatBlt пишется в EMF как EMR_BITBLT без исходного битмапа
        private void HandleBitBlt(IntPtr pRec)
        {
            int xDest = Marshal.ReadInt32(pRec, 24);
            int yDest = Marshal.ReadInt32(pRec, 28);
            int cx = Marshal.ReadInt32(pRec, 32);
            int cy = Marshal.ReadInt32(pRec, 36);
            uint rop = ReadU32(pRec, 40);
            uint fill;
            if (rop == 0x00000042) fill = 0x000000;          // BLACKNESS
            else if (rop == 0x00FF0062) fill = 0xFFFFFF;     // WHITENESS
            else if (rop == 0x00F00021) fill = _brushColor;  // PATCOPY: текущая кисть
            else return;
            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                "{0} {1} {2} {3} re {4} rg f\n", xDest, yDest, cx, cy, ColorToPdf(fill));
        }
        private void HandleBezier16(IntPtr pRec)
        {
            // EMR_POLYBEZIER16: iType, nSize, rclBounds(16), cpts(4), aptl[]
            uint count = ReadU32(pRec, 24);
            int pointsOff = 28;
            // Первая точка — moveto, далее по 3 точки — кривая
            short x0 = Marshal.ReadInt16(pRec, pointsOff + 0);
            short y0 = Marshal.ReadInt16(pRec, pointsOff + 2);
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} m\n", (double)x0, (double)y0);
            for (uint i = 1; i + 2 < count; i += 3)
            {
                short c1x = Marshal.ReadInt16(pRec, pointsOff + (int)(i * 4));
                short c1y = Marshal.ReadInt16(pRec, pointsOff + (int)(i * 4) + 2);
                short c2x = Marshal.ReadInt16(pRec, pointsOff + (int)((i + 1) * 4));
                short c2y = Marshal.ReadInt16(pRec, pointsOff + (int)((i + 1) * 4) + 2);
                short ex = Marshal.ReadInt16(pRec, pointsOff + (int)((i + 2) * 4));
                short ey = Marshal.ReadInt16(pRec, pointsOff + (int)((i + 2) * 4) + 2);
                _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                    "{0:0.###} {1:0.###} {2:0.###} {3:0.###} {4:0.###} {5:0.###} c\n",
                    (double)c1x, (double)c1y, (double)c2x, (double)c2y, (double)ex, (double)ey);
            }
            Stroke();
        }

        private void HandleMoveTo(IntPtr pRec)
        {
            // EMR_MOVETOEX: iType, nSize, ptl (POINTL)
            _curPosX = Marshal.ReadInt32(pRec, 8);
            _curPosY = Marshal.ReadInt32(pRec, 12);
            //_page.Content.AppendFormat(CultureInfo.InvariantCulture,
            //    "{0:0.###} {1:0.###} m\n", (double)_curPosX, (double)_curPosY);
        }

        private void HandleLineTo(IntPtr pRec)
        {
            // EMR_LINETO: iType, nSize, ptl (POINTL). Линия рисуется от текущей позиции.
            int x = Marshal.ReadInt32(pRec, 8);
            int y = Marshal.ReadInt32(pRec, 12);
            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                "{0} {1} m {2} {3} l ", _curPosX, _curPosY, x, y);
            Stroke();
            _curPosX = x; _curPosY = y;
        }

        // -------------------- Текст --------------------

        private void HandleText(IntPtr hdc, IntPtr pRec, uint size)
        {
            int refX = Marshal.ReadInt32(pRec, 36);
            int refY = Marshal.ReadInt32(pRec, 40);
            int nChars = Marshal.ReadInt32(pRec, 44);
            uint offString = ReadU32(pRec, 48);
            // offDx не нужен — мы не делаем попиксельный кернинг, PDF Tj сам шагает по Widths.

            if (nChars <= 0 || offString == 0 || offString + (uint)nChars * 2 > size) return;

            // Читаем строку (UTF-16LE) из записи по смещению offString
            string text = ReadUtf16(pRec, (int)offString, nChars);
            // Текущий шрифт: bold/regular
            bool bold = _fontBold;
            int fi = (_fontBold ? 1 : 0) | (_fontItalic ? 2 : 0);
            var fnt = _fonts[fi];
            string fontRef = "/F" + (fi + 1);
            double skew = (_fontItalic && _synthItalic[fi]) ? 0.2126 : 0;   // tan(12°)
            // Размер в EMF-единицах: CTM (72/Dpi) сам переведёт в пункты
            double sizePx = _fontEmHeightInPx > 0 ? _fontEmHeightInPx : 16;

            var hex = new StringBuilder(text.Length * 4);
            double wsum = 0;
            foreach (char ch in text)
            {
                ushort gid;
                fnt.Cmap.TryGetValue(ch, out gid);
                if (gid != 0 && !fnt.UsedGlyphs.ContainsKey(gid)) fnt.UsedGlyphs[gid] = ch;
                hex.Append(gid.ToString("X4"));
                wsum += gid < fnt.GlyphWidths.Length ? fnt.GlyphWidths[gid] : 1000;
            }


            // ptlReference: по умолчанию (TA_TOP) это ВЕРХ текста, не baseline
            double baseY = refY;
            uint vAlign = _textAlign & 0x18;          // 0=TA_TOP, 8=TA_BOTTOM, 24=TA_BASELINE
            double ascPx = sizePx * (fnt.WinAscent > 0 ? fnt.WinAscent / 1000.0 : 0.85);
            if (vAlign == 0) 
                baseY = refY + ascPx;    // top -> baseline
            else if (vAlign == 8)
                baseY = refY;       // bottom ~ baseline (приближение)

            // refY в EMF — обычно baseline. В user space после базового cm Y идёт ВНИЗ.
            // PDF-текст "растёт вверх" от baseline в text-space, что после CTM
            // соответствовало бы движению ВНИЗ по странице. Поэтому используем Tm с
            // отражением по Y: 1 0 0 -1 refX refY — baseline в (refX, refY),
            // текст растёт ВВЕРХ по странице.
            _page.Content.Append("BT\n");
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} {1:0.##} Tf\n", fontRef, sizePx);
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} rg\n", ColorToPdf(_textColor));
            //_page.Content.AppendFormat(CultureInfo.InvariantCulture, "1 0 0 -1 {0:0.####} {1:0.####} Tm\n", (double)refX, baseY);
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "1 0 {0:0.####} -1 {1:0.###} {2:0.###} Tm\n", skew, (double)refX, baseY);
            _page.Content.Append("<").Append(hex).Append("> Tj\nET\n");

            if (_fontUnderline || _fontStrike)
            {
                double uth = (fnt.UnderlineThickness > 0 ? fnt.UnderlineThickness / 1000.0 : 0.07) * sizePx;
                string col = ColorToPdf(_textColor);
                if (_fontUnderline)
                {
                    // y вниз: ниже baseline = плюс
                    double uy = baseY + (fnt.UnderlinePos != 0 ? -fnt.UnderlinePos / 1000.0 : 0.12) * sizePx;
                    _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                        "{0} rg {1:0.###} {2:0.###} {3:0.###} {4:0.###} re f\n", col, (double)refX, uy, wsum, uth);
                }
                if (_fontStrike)
                {
                    double sy = baseY - 0.3 * sizePx;
                    _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                        "{0} rg {1:0.###} {2:0.###} {3:0.###} {4:0.###} re f\n", col, (double)refX, sy, wsum, uth);
                }
            }
        }

        private void HandleStretchDibits(IntPtr pRec, uint size)
        {
            // EMR_STRETCHDIBITS:
            //   iType(4) nSize(4) rclBounds(16) xDest(4) yDest(4) xSrc(4) ySrc(4) cxSrc(4) cySrc(4)
            //   offBmiSrc(4) cbBmiSrc(4) offBitsSrc(4) cbBitsSrc(4) iUsageSrc(4) dwRop(4) cxDest(4) cyDest(4)
            int xDest = Marshal.ReadInt32(pRec, 24);
            int yDest = Marshal.ReadInt32(pRec, 28);
            int xSrc = Marshal.ReadInt32(pRec, 32);
            int ySrc = Marshal.ReadInt32(pRec, 36);
            int cxSrc = Marshal.ReadInt32(pRec, 40);
            int cySrc = Marshal.ReadInt32(pRec, 44);
            uint offBmi = ReadU32(pRec, 48);
            uint cbBmi = ReadU32(pRec, 52);
            uint offBits = ReadU32(pRec, 56);
            uint cbBits = ReadU32(pRec, 60);
            uint iUsage = ReadU32(pRec, 64);
            uint dwRop = ReadU32(pRec, 68);
            int cxDest = Marshal.ReadInt32(pRec, 72);
            int cyDest = Marshal.ReadInt32(pRec, 76);

            if (cbBmi == 0 || cbBits == 0) return;
            if (offBmi + cbBmi > size || offBits + cbBits > size) return;

            // Считываем BMI и биты
            byte[] bmi = new byte[cbBmi];
            Marshal.Copy(IntPtr.Add(pRec, (int)offBmi), bmi, 0, (int)cbBmi);
            byte[] bits = new byte[cbBits];
            Marshal.Copy(IntPtr.Add(pRec, (int)offBits), bits, 0, (int)cbBits);

            // Конвертируем DIB в RGB-байты (24-bit несжатый)
            var img = DibToRgb(bmi, bits, xSrc, ySrc, cxSrc, cySrc);
            if (img == null || img.Data == null) return;

            // Сжимаем через deflate и добавляем как XObject
            byte[] compressed = PdfZlib.Deflate(img.Data);
            int imgId = _pdf.AddImage(compressed, img.Width, img.Height, false);
            int localIdx = _imageIdsOnPage.Count + 1;
            _imageIdsOnPage.Add(imgId);
            _page.ImageRefs.Add(imgId);

            // Рисуем XObject. Координаты: xDest, yDest — левый ВЕРХНИЙ угол в EMF.
            // После отражения Y-пикселей PDF-страницы yDest становится нижней границей.
            // Изображение размещаем как /Im<n> Do с cm(width 0 0 height x y).
            double x = xDest;
            double y = yDest;   // уже отражён в верх-низ через базовый cm
            double wpt = cxDest;
            double hpt = cyDest;
            _page.Content.Append("q\n");
            _page.Content.AppendFormat(CultureInfo.InvariantCulture,
                "{0:0.###} 0 0 {1:0.###} {2:0.###} {3:0.###} cm\n", wpt, hpt, x, y);
            _page.Content.AppendFormat(CultureInfo.InvariantCulture, "/Im{0} Do\nQ\n", localIdx);
        }

        // -------------------- Вспомогательные --------------------

        private void Fill() => _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} rg f\n", ColorToPdf(_brushColor));
        private void Stroke() => _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} RG {1:0.##} w S\n", ColorToPdf(_penColor), _penWidth);
        private void FillStroke(bool doFill, bool doStroke)
        {
            if (doFill && doStroke)
                _page.Content.AppendFormat(CultureInfo.InvariantCulture, "{0} rg {1} RG B\n",
                    ColorToPdf(_brushColor), ColorToPdf(_penColor));
            else if (doFill) Fill();
            else if (doStroke) Stroke();
        }

        private static string ColorToPdf(uint color)
        {
            double r = (color & 0xFF) / 255.0;
            double g = ((color >> 8) & 0xFF) / 255.0;
            double b = ((color >> 16) & 0xFF) / 255.0;
            return r.ToString("0.###", CultureInfo.InvariantCulture) + " " +
                   g.ToString("0.###", CultureInfo.InvariantCulture) + " " +
                   b.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static float ReadFloat(IntPtr p, int off)
        {
            int bits = Marshal.ReadInt32(p, off);
            // Без unsafe: int -> float через BitConverter.
            byte[] b = BitConverter.GetBytes(bits);
            return BitConverter.ToSingle(b, 0);
        }
        private static uint ReadU32(IntPtr p, int off) => (uint)Marshal.ReadInt32(p, off);

        private static string ReadUtf16(IntPtr p, int off, int count)
        {
            var sb = new StringBuilder(count);
            for (int i = 0; i < count; i++)
                sb.Append((char)Marshal.ReadInt16(p, off + i * 2));
            return sb.ToString();
        }

        private static string EscapePdfString(string s)
        {
            // Строки PDF: экранируем ( ) \ и удаляем непечатные управляющие.
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                if (ch == '(' || ch == ')' || ch == '\\')
                    sb.Append('\\').Append(ch);
                else if (ch < 32)
                    sb.Append(' ');
                else
                    sb.Append(ch);
            }
            return sb.ToString();
        }

        // Unicode → CP1251 byte (для помещения в content stream).
        // Возвращает -1, если символ в CP1251 не представлен.
        private static int UnicodeToCp1251(char ch)
        {
            if (ch < 0x80) return ch;
            // Основной русский диапазон
            if (ch >= '\u0410' && ch <= '\u044F') return 0xC0 + (ch - '\u0410');
            // Кириллические дополнения (расширенный русский + украинский/белорусский)
            switch (ch)
            {
                case '\u0401': return 0xA8; // Ё
                case '\u0451': return 0xB8; // ё
                case '\u040E': return 0xA1; // Ў (бел.)
                case '\u045E': return 0xA2; // ў
                case '\u0404': return 0xAA; // Є (укр.)
                case '\u0454': return 0xBA; // є
                case '\u0406': return 0xB2; // І
                case '\u0456': return 0xB3; // і
                case '\u0407': return 0xAF; // Ї
                case '\u0457': return 0xBF; // ї
                case '\u0402': return 0xA3; // Ђ (серб.)
                case '\u0452': return 0xA4; // ђ
                case '\u0405': return 0xA5; // Ѕ (макед.)
                case '\u0455': return 0xB5; // ѕ  (note: в CP1251 0xB5 = ѕ, не µ)
                case '\u0408': return 0xA6; // Ј
                case '\u0458': return 0xA6; // ј (конфликт — оба на 0xA6 в разных источниках)
                case '\u040C': return 0xA7; // Ќ
                case '\u0409': return 0xA6; // Љ (то же конфликтное место)
                case '\u040A': return 0xA6; // Њ
                case '\u040B': return 0xA6; // Ћ
                case '\u0403': return 0xA3; // Ѓ
            }
            // Latin-1 chars 0xA0..0xFF: в CP1251 большинство совпадает с Latin-1
            // (это пунктуация: ©, ®, «, », №, ° и т.д.). Возвращаем как есть.
            if (ch >= 0xA0 && ch <= 0xFF) return (int)ch;
            return -1;
        }

        // ---------- DIB -> RGB ----------

        private static RgbImage DibToRgb(byte[] bmi, byte[] bits, int xSrc, int ySrc, int cxSrc, int cySrc)
        {
            try
            {
                if (bmi.Length < 40) return null;
                uint headerSize = BitConverter.ToUInt32(bmi, 0);
                int width = BitConverter.ToInt32(bmi, 4);
                int height = BitConverter.ToInt32(bmi, 8);
                ushort bpp = BitConverter.ToUInt16(bmi, 14);
                uint compression = BitConverter.ToUInt32(bmi, 16);

                if (width <= 0 || height == 0) return null;
                bool bottomUp = height > 0;
                int absH = Math.Abs(height);

                if (compression != 0)
                {
                    // RLE / JPEG / PNG — fallback: рендерить через System.Drawing
                    return DibToRgbViaGdi(bmi, bits, xSrc, ySrc, cxSrc, cySrc);
                }

                // Поддержка 24/32 bpp. 1/4/8 — палитра, конвертируем.
                int stride = ((width * bpp + 31) / 32) * 4;
                int palCount = 0;
                int[] palette = null;
                if (bpp <= 8)
                {
                    palCount = (int)((headerSize == 12 ? 0 : BitConverter.ToUInt32(bmi, 32)));
                    if (palCount == 0) palCount = 1 << bpp;
                    int palEntrySize = (headerSize == 12) ? 3 : 4;
                    palette = new int[palCount];
                    int palOff = (int)headerSize;
                    for (int i = 0; i < palCount; i++)
                    {
                        int b = bmi[palOff + i * palEntrySize + 0];
                        int g = bmi[palOff + i * palEntrySize + 1];
                        int r = bmi[palOff + i * palEntrySize + 2];
                        palette[i] = (r << 16) | (g << 8) | b;
                    }
                }

                // Итоговый размер: cxSrc * cySrc * 3
                int outW = cxSrc, outH = cySrc;
                byte[] rgb = new byte[outW * outH * 3];

                // Коэффициенты масштабирования (src -> dst)
                double sx = (double)cxSrc / Math.Max(1, width);
                double sy = (double)cySrc / Math.Max(1, absH);

                for (int dy = 0; dy < outH; dy++)
                {
                    int srcY = (int)(dy / sy);
                    if (!bottomUp) srcY = absH - 1 - srcY;
                    int rowOff = srcY * stride;
                    for (int dx = 0; dx < outW; dx++)
                    {
                        int srcX = (int)(dx / sx);
                        int r = 0, g = 0, b = 0;
                        if (bpp == 24)
                        {
                            int off = rowOff + srcX * 3;
                            if (off + 2 < bits.Length)
                            { b = bits[off]; g = bits[off + 1]; r = bits[off + 2]; }
                        }
                        else if (bpp == 32)
                        {
                            int off = rowOff + srcX * 4;
                            if (off + 2 < bits.Length)
                            { b = bits[off]; g = bits[off + 1]; r = bits[off + 2]; }
                        }
                        else if (bpp == 8)
                        {
                            int off = rowOff + srcX;
                            if (off < bits.Length)
                            {
                                int c = palette[bits[off]];
                                r = (c >> 16) & 0xFF; g = (c >> 8) & 0xFF; b = c & 0xFF;
                            }
                        }
                        else if (bpp == 4)
                        {
                            int off = rowOff + (srcX >> 1);
                            if (off < bits.Length)
                            {
                                int c = palette[(bits[off] >> ((1 - (srcX & 1)) * 4)) & 0x0F];
                                r = (c >> 16) & 0xFF; g = (c >> 8) & 0xFF; b = c & 0xFF;
                            }
                        }
                        else if (bpp == 1)
                        {
                            int off = rowOff + (srcX >> 3);
                            if (off < bits.Length)
                            {
                                int c = palette[(bits[off] >> (7 - (srcX & 7))) & 1];
                                r = (c >> 16) & 0xFF; g = (c >> 8) & 0xFF; b = c & 0xFF;
                            }
                        }
                        int o = (dy * outW + dx) * 3;
                        rgb[o] = (byte)r; rgb[o + 1] = (byte)g; rgb[o + 2] = (byte)b;
                    }
                }
                return new RgbImage { Data = rgb, Width = outW, Height = outH };
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DibToRgb: " + ex.Message);
                return null;
            }
        }

        // Fallback для сжатых DIB (RLE4/RLE8/JPEG/PNG): используем System.Drawing.
        private static RgbImage DibToRgbViaGdi(byte[] bmi, byte[] bits, int xSrc, int ySrc, int cxSrc, int cySrc)
        {
            try
            {
                // Собираем полный DIB-буфер (BMI + bits) и грузим через MemoryImageSource.
                // Самый простой путь — записать BMP-файл во временный MemoryStream и открыть через Image.
                if (bmi.Length < 40) return null;

                // BMP file header (14) + DIB
                using (var ms = new MemoryStream())
                using (var bw = new BinaryWriter(ms, Encoding.ASCII, leaveOpen: true))
                {
                    bw.Write((byte)'B'); bw.Write((byte)'M');
                    bw.Write((uint)(14 + bmi.Length + bits.Length));
                    bw.Write((ushort)0); bw.Write((ushort)0);
                    bw.Write((uint)(14 + bmi.Length));
                    bw.Write(bmi, 0, bmi.Length);
                    bw.Write(bits, 0, bits.Length);

                    ms.Position = 0;
                    using (var img = Image.FromStream(ms))
                    {
                        int outW = cxSrc <= 0 ? img.Width : cxSrc;
                        int outH = cySrc <= 0 ? img.Height : cySrc;
                        using (var bmp = new Bitmap(outW, outH, PixelFormat.Format24bppRgb))
                        using (var g = Graphics.FromImage(bmp))
                        {
                            g.DrawImage(img, 0, 0, outW, outH);
                            var bd = bmp.LockBits(new Rectangle(0, 0, outW, outH), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                            try
                            {
                                int stride = bd.Stride;
                                byte[] rgb = new byte[outW * outH * 3];
                                for (int y = 0; y < outH; y++)
                                {
                                    Marshal.Copy(new IntPtr(bd.Scan0.ToInt64() + y * stride), rgb, y * outW * 3, outW * 3);
                                }
                                return new RgbImage { Data = rgb, Width = outW, Height = outH };
                            }
                            finally
                            {
                                bmp.UnlockBits(bd);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DibToRgbViaGdi: " + ex.Message);
                return null;
            }
        }
    }
}