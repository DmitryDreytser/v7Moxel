// EmfVectorPdfWriter.cs
// Полноценный ВЕКТОРНЫЙ конвертер EMF -> PDF.
// Распарсивает записи GDI (GDI-команды) из EMF и транслирует их в PDF Content Stream.
// Текст, линии, прямоугольники — идеальные вектора. Изображения — встроенные XObject'ы.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Moxel
{
    public static class EmfVectorPdfWriter2
    {
        #region P/Invoke и структуры EMF

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr SetEnhMetaFileBits(uint nSize, byte[] lpData);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetEnhMetaFileHeader(IntPtr hEmf, uint nSize, out ENHMETAHEADER hdr);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool EnumEnhMetaFile(IntPtr hdc, IntPtr hemf, EnumEnhMetaFileProc lpProc, IntPtr lpData, IntPtr lpRect);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr GetCurrentObject(IntPtr hdc, uint objectType);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern int GetObject(IntPtr hgdiobj, int cbBuffer, ref LOGFONT lf);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern int GetObject(IntPtr hgdiobj, int cbBuffer, ref LOGBRUSH lb);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern int GetObject(IntPtr hgdiobj, int cbBuffer, ref EXTLOGPEN lp);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetObjectType(IntPtr hgdiobj);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteEnhMetaFile(IntPtr hEmf);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateDIBitmap(IntPtr hdc, IntPtr pbmih, uint flInit, IntPtr pjInit, IntPtr pbmi, uint iUsage);

        private const uint OBJ_PEN = 1;
        private const uint OBJ_BRUSH = 2;
        private const uint OBJ_FONT = 6;
        private const uint ENHMETA_SIGNATURE = 0x464D4520;

        [StructLayout(LayoutKind.Sequential)]
        private struct ENHMETARECORD
        {
            public int iType;
            public uint nSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ENHMETAHEADER
        {
            public int iType;
            public uint nSize;
            public RECT rclBounds;
            public RECT rclFrame;
            public uint dSignature;
            public uint nVersion;
            public uint nBytes;
            public uint nRecords;
            public ushort nHandles;
            public ushort sReserved;
            public uint nDescription;
            public uint offDescription;
            public uint nPalEntries;
            public int szlDevice1;
            public int szlDevice2;
            public int szlMillimeters1;
            public int szlMillimeters2;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LOGFONT
        {
            public int lfHeight;
            public int lfWidth;
            public int lfEscapement;
            public int lfOrientation;
            public int lfWeight;
            public byte lfItalic;
            public byte lfUnderline;
            public byte lfStrikeOut;
            public byte lfCharSet;
            public byte lfOutPrecision;
            public byte lfClipPrecision;
            public byte lfQuality;
            public byte lfPitchAndFamily;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string lfFaceName;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LOGBRUSH
        {
            public uint lbStyle;
            public uint lbColor;
            public uint lbHatch;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct EXTLOGPEN
        {
            public uint elpPenStyle;
            public uint elpWidth;
            public uint elpBrushStyle;
            public uint elpColor;
            public uint elpHatch;
            public uint elpNumEntries;
            public uint elpStyleEntry;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct EMRSELECTOBJECT
        {
            public int iType;
            public uint nSize;
            public uint ihObject;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct EMRRECTANGLE
        {
            public int iType;
            public uint nSize;
            public RECT rclBox;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int left, top, right, bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int x, y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct EMRTEXT
        {
            public POINT ptlReference;
            public uint nChars;
            public uint offString;
            public uint fOptions;
            public RECT rcl;
            public uint offDx;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct EMREXTTEXTOUTW
        {
            public int iType;
            public uint nSize;
            public RECT rclBounds;
            public int iGraphicsMode;
            public float exScale;
            public float eyScale;
            public EMRTEXT emrText;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct EMRSETTEXTCOLOR
        {
            public int iType;
            public uint nSize;
            public uint crColor;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct EMRSTRETCHDIBITS
        {
            public int iType;
            public uint nSize;
            public RECT rclBounds;
            public int xDest;
            public int yDest;
            public int xSrc;
            public int ySrc;
            public int cxSrc;
            public int cySrc;
            public uint offBmiSrc;
            public uint cbBmiSrc;
            public uint offBitsSrc;
            public uint cbBitsSrc;
            public uint iUsageSrc;
            public uint dwRop;
            public int cxDest;
            public int cyDest;
        }

        #endregion

        private enum EmrRecordType : int
        {
            EMR_HEADER = 1,
            EMR_SETTEXTCOLOR = 10,
            EMR_SELECTOBJECT = 37,
            EMR_CREATEBRUSHINDIRECT = 39,
            EMR_DELETEOBJECT = 40,
            EMR_EXTCREATEPEN = 97,
            EMR_RECTANGLE = 43,
            EMR_POLYLINE16 = 87,
            EMR_POLYGON16 = 82,
            EMR_SETWORLDTRANSFORM = 23,
            EMR_EXTTEXTOUTW = 84,
            EMR_STRETCHDIBITS = 81,
            EMR_EOF = 109
        }

        private delegate bool EnumEnhMetaFileProc(IntPtr hdc, IntPtr ht, IntPtr mr, uint hHandles, IntPtr data);

        public static bool Save(IList<byte[]> emfPages, string fileName, EmfPdfSettings settings = null)
        {
            if (emfPages == null || emfPages.Count == 0 || string.IsNullOrEmpty(fileName)) return false;

            try
            {
                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, 256 * 1024))
                {
                    // Используем словарь для надежного хранения смещений объектов
                    var offs = new Dictionary<int, long>();
                    int nextObjId = 1;

                    Action<string> W = s => { byte[] b = Encoding.ASCII.GetBytes(s); fs.Write(b, 0, b.Length); };
                    Action<byte[]> WB = b => fs.Write(b, 0, b.Length);
                    Action<int, Action> Obj = (id, body) => { offs[id] = fs.Position; W(id + " 0 obj\n"); body(); W("\nendobj\n"); };

                    W("%PDF-1.4\n");
                    WB(new byte[] { 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A });

                    int catalogId = nextObjId++;
                    int pagesId = nextObjId++;
                    int fontHelvId = nextObjId++;
                    int fontHelvBoldId = nextObjId++;

                    Obj(fontHelvId, () => W("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"));
                    Obj(fontHelvBoldId, () => W("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"));

                    Obj(catalogId, () => W($"<< /Type /Catalog /Pages {pagesId} 0 R >>"));

                    var pageIds = new List<int>();
                    for (int i = 0; i < emfPages.Count; i++)
                        pageIds.Add(nextObjId++); // ID для Page

                    var contentIds = new List<int>();
                    for (int i = 0; i < emfPages.Count; i++)
                        contentIds.Add(nextObjId++); // ID для Content Stream

                    Obj(pagesId, () =>
                    {
                        var kids = new StringBuilder();
                        foreach (var pid in pageIds) kids.Append(pid + " 0 R ");
                        W($"<< /Type /Pages /Count {emfPages.Count} /Kids [ {kids}] >>");
                    });

                    for (int i = 0; i < emfPages.Count; i++)
                    {
                        IntPtr hmf = SetEnhMetaFileBits((uint)emfPages[i].Length, emfPages[i]);
                        if (hmf == IntPtr.Zero) continue;

                        try
                        {
                            if (GetEnhMetaFileHeader(hmf, (uint)Marshal.SizeOf(typeof(ENHMETAHEADER)), out var hdr) == 0 || hdr.dSignature != ENHMETA_SIGNATURE)
                                continue;

                            double mmW = Math.Max(1, hdr.rclFrame.right - hdr.rclFrame.left) / 100.0;
                            double mmH = Math.Max(1, hdr.rclFrame.bottom - hdr.rclFrame.top) / 100.0;
                            double mediaWpt = mmW * 72.0 / 25.4;
                            double mediaHpt = mmH * 72.0 / 25.4;

                            int pageId = pageIds[i];
                            int contId = contentIds[i];

                            var engine = new PageEngine(hmf, settings ?? new EmfPdfSettings(), nextObjId, fs, offs);
                            engine.Process();
                            nextObjId = engine.NextObjId;

                            Obj(pageId, () =>
                                W($"<< /Type /Page /Parent {pagesId} 0 R /MediaBox [0 0 {Inv(mediaWpt)} {Inv(mediaHpt)}] " +
                                  $"/Resources << /Font << /F1 {fontHelvId} 0 R /F2 {fontHelvBoldId} 0 R >> {engine.GetXObjects()} >> " +
                                  $"/Contents {contId} 0 R >>"));

                            byte[] content = Zlib(Encoding.GetEncoding(1251).GetBytes(engine.GetContent()));
                            Obj(contId, () => {
                                W($"<< /Length {content.Length} /Filter /FlateDecode >>\nstream\n");
                                WB(content);
                                W("\nendstream");
                            });
                        }
                        finally
                        {
                            DeleteEnhMetaFile(hmf);
                        }
                    }

                    long xrefPos = fs.Position;
                    W($"xref\n0 {nextObjId}\n");
                    W("0000000000 65535 f \n");
                    for (int id = 1; id < nextObjId; id++)
                    {
                        if (offs.ContainsKey(id))
                            W($"{offs[id]:D10} 00000 n \n");
                        else
                            W("0000000000 00000 f \n");
                    }
                    W($"trailer\n<< /Size {nextObjId} /Root {catalogId} 0 R >>\nstartxref\n{xrefPos}\n%%EOF");
                }
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("EmfVectorPdfWriter: " + ex);
                return false;
            }
        }

        private class PageEngine
        {
            private IntPtr _hmf;
            private EmfPdfSettings _settings;
            private StringBuilder _content = new StringBuilder();
            private Dictionary<string, int> _images = new Dictionary<string, int>();
            private float _scale;
            private int _nextObjId;
            private Dictionary<int, long> _offs;
            private FileStream _fs;
            private int _pageHeightPx;

            private uint _curTextColor = 0x000000;
            private float _curPenWidth = 0.5f;
            private uint _curPenColor = 0x000000;
            private uint _curBrushColor = 0xFFFFFF;
            private bool _boldFont = false;

            public int NextObjId => _nextObjId;

            public PageEngine(IntPtr hmf, EmfPdfSettings settings, int startObjId, FileStream fs, Dictionary<int, long> offs)
            {
                _hmf = hmf;
                _settings = settings;
                _scale = 72.0f / (_settings.Dpi <= 0 ? 96 : _settings.Dpi);
                _nextObjId = startObjId;
                _offs = offs;
                _fs = fs;

                // Получаем высоту в пикселях для инверсии оси Y
                ENHMETAHEADER hdr;
                GetEnhMetaFileHeader(hmf, (uint)Marshal.SizeOf(typeof(ENHMETAHEADER)), out hdr);
                _pageHeightPx = hdr.rclBounds.bottom - hdr.rclBounds.top;

                // Матрица: масштабирование по X, масштабирование и инверсия по Y, сдвиг по Y
                _content.AppendLine("q");
                _content.AppendFormat(System.Globalization.CultureInfo.InvariantCulture,
                    "{0:F5} 0 0 -{0:F5} 0 {1:F2} cm\n", _scale, _pageHeightPx * _scale);
                _content.AppendLine("0.5 w");
                _content.AppendLine("0 0 0 RG");
                _content.AppendLine("1 1 1 rg");
            }

            public string GetContent()
            {
                _content.AppendLine("Q");
                return _content.ToString();
            }

            public string GetXObjects()
            {
                if (_images.Count == 0) return "";
                var sb = new StringBuilder("/XObject << ");
                foreach (var kvp in _images)
                    sb.Append($"/{kvp.Key} {kvp.Value} 0 R ");
                sb.Append(">>");
                return sb.ToString();
            }

            public void Process()
            {
                GCHandle handle = GCHandle.Alloc(this);
                try
                {
                    EnumEnhMetaFile(IntPtr.Zero, _hmf, Callback, GCHandle.ToIntPtr(handle), IntPtr.Zero);
                }
                finally
                {
                    handle.Free();
                }
            }

            private bool Callback(IntPtr hdc, IntPtr ht, IntPtr mr, uint hHandles, IntPtr data)
            {
                int iType = Marshal.ReadInt32(mr);
                ProcessRecord(hdc, ht, mr, iType);
                return true;
            }

            private void ProcessRecord(IntPtr hdc, IntPtr ht, IntPtr mr, int iType)
            {
                switch ((EmrRecordType)iType)
                {
                    case EmrRecordType.EMR_SETTEXTCOLOR:
                        _curTextColor = (uint)Marshal.ReadInt32(mr, 8);
                        break;

                    case EmrRecordType.EMR_CREATEBRUSHINDIRECT:
                        // 8 (emr) + 4 (ihBrush) + 4 (lbStyle) = 16. Цвет начинается с 16 байта
                        _curBrushColor = (uint)Marshal.ReadInt32(mr, 16);
                        break;

                    case EmrRecordType.EMR_SELECTOBJECT:
                        SelectObject(hdc, ht, mr);
                        break;

                    case EmrRecordType.EMR_RECTANGLE:
                        DrawRectangle(mr);
                        break;

                    case EmrRecordType.EMR_POLYLINE16:
                        DrawPolyline(mr, false);
                        break;

                    case EmrRecordType.EMR_POLYGON16:
                        DrawPolyline(mr, true);
                        break;

                    case EmrRecordType.EMR_EXTTEXTOUTW:
                        DrawText(hdc, mr);
                        break;

                    case EmrRecordType.EMR_STRETCHDIBITS:
                        DrawImage(mr);
                        break;
                }
            }

            private void SelectObject(IntPtr hdc, IntPtr ht, IntPtr mr)
            {
                uint ihObject = (uint)Marshal.ReadInt32(mr, 8);
                IntPtr hObj = Marshal.ReadIntPtr(ht, (int)(ihObject * IntPtr.Size));
                uint type = GetObjectType(hObj);

                if (type == OBJ_PEN)
                {
                    var pen = new EXTLOGPEN();
                    GetObject(hObj, Marshal.SizeOf(typeof(EXTLOGPEN)), ref pen);
                    _curPenColor = pen.elpColor;
                    _curPenWidth = Math.Max(0.1f, pen.elpWidth * 0.5f);
                    _content.AppendLine($"{ColorToPdf(_curPenColor)} RG");
                    _content.AppendLine($"{_curPenWidth.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)} w");
                }
                else if (type == OBJ_BRUSH)
                {
                    var lb = new LOGBRUSH();
                    GetObject(hObj, Marshal.SizeOf(typeof(LOGBRUSH)), ref lb);
                    _curBrushColor = lb.lbColor;
                }
            }

            private void DrawRectangle(IntPtr mr)
            {
                var rect = Marshal.PtrToStructure<EMRRECTANGLE>(mr);
                float w = rect.rclBox.right - rect.rclBox.left;
                float h = rect.rclBox.bottom - rect.rclBox.top;

                // В PDF начало координат прямоугольника - нижний левый угол.
                // С учетом матрицы инверсии, передаем rclBox.bottom как Y
                _content.AppendLine($"{rect.rclBox.left} {rect.rclBox.bottom} {w} {h} re");
                _content.AppendLine($"{ColorToPdf(_curBrushColor)} rg f");
                _content.AppendLine($"{ColorToPdf(_curPenColor)} RG S");
            }

            private void DrawPolyline(IntPtr mr, bool close)
            {
                // EMR_POLYLINE16/POLYGON16: 8 (emr) + 16 (bounds) + 4 (count)
                int cptl = Marshal.ReadInt32(mr, 24);
                if (cptl == 0) return;

                _content.AppendLine("newpath");
                for (int i = 0; i < cptl; i++)
                {
                    // Смещение: 28 + i * 4 (каждая точка 2 short'a = 4 байта)
                    int offset = 28 + i * 4;
                    short x = Marshal.ReadInt16(mr, offset);
                    short y = Marshal.ReadInt16(mr, offset + 2);

                    if (i == 0)
                        _content.AppendLine($"{x} {y} m");
                    else
                        _content.AppendLine($"{x} {y} l");
                }
                if (close) _content.AppendLine("h");
                _content.AppendLine($"{ColorToPdf(_curPenColor)} RG S");
            }

            private void DrawText(IntPtr hdc, IntPtr mr)
            {
                var eto = Marshal.PtrToStructure<EMREXTTEXTOUTW>(mr);
                if (eto.emrText.nChars == 0) return;

                IntPtr pStr = IntPtr.Add(mr, (int)eto.emrText.offString);
                string text = Marshal.PtrToStringUni(pStr, (int)eto.emrText.nChars);

                IntPtr hFont = GetCurrentObject(hdc, OBJ_FONT);
                var lf = new LOGFONT();
                GetObject(hFont, Marshal.SizeOf(typeof(LOGFONT)), ref lf);

                float size = Math.Abs(lf.lfHeight) * 72.0f / (_settings.Dpi <= 0 ? 96 : _settings.Dpi);
                size = Math.Max(1, size);

                _boldFont = lf.lfWeight >= 700;
                string fontName = _boldFont ? "F2" : "F1";
                float angle = lf.lfEscapement / 10.0f;

                _content.AppendLine("BT");
                _content.AppendLine($"{fontName} {size.ToString("F1", System.Globalization.CultureInfo.InvariantCulture)} Tf");
                _content.AppendLine($"{ColorToPdf(_curTextColor)} rg");

                // Точка вставки. В 1С обычно ptlReference - это baseline (или top), берем с небольшим запасом снизу
                float x = eto.emrText.ptlReference.x;
                float y = eto.emrText.ptlReference.y - size * 0.8f; // Корректировка базовой линии

                if (angle == 0)
                {
                    _content.AppendLine($"{x.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)} {y.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)} Td");
                }
                else
                {
                    double rad = angle * Math.PI / 180.0;
                    float cos = (float)Math.Cos(rad);
                    float sin = (float)Math.Sin(rad);
                    _content.AppendLine($"{cos.ToString("F5", System.Globalization.CultureInfo.InvariantCulture)} {sin.ToString("F5", System.Globalization.CultureInfo.InvariantCulture)} {(-sin).ToString("F5", System.Globalization.CultureInfo.InvariantCulture)} {cos.ToString("F5", System.Globalization.CultureInfo.InvariantCulture)} {x.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)} {y.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)} Tm");
                }

                // Экранируем и кодируем в 1251 (WinAnsi)
                byte[] win1251Bytes = Encoding.GetEncoding(1251).GetBytes(text);
                StringBuilder hexStr = new StringBuilder();
                foreach (byte b in win1251Bytes) hexStr.AppendFormat("{0:X2}", b);

                _content.AppendLine($"<{hexStr}> Tj");
                _content.AppendLine("ET");
            }

            private void DrawImage(IntPtr mr)
            {
                var sdb = Marshal.PtrToStructure<EMRSTRETCHDIBITS>(mr);

                IntPtr pBmi = IntPtr.Add(mr, (int)sdb.offBmiSrc);
                IntPtr pBits = IntPtr.Add(mr, (int)sdb.offBitsSrc);

                try
                {
                    IntPtr hBmp = CreateDIBitmap(IntPtr.Zero, pBmi, 0x00000004 /*CB_INIT*/, pBits, pBmi, 0 /*DIB_RGB_COLORS*/);
                    if (hBmp == IntPtr.Zero) return;

                    using (Bitmap bmp = Image.FromHbitmap(hBmp))
                    {
                        int imgId = _nextObjId++;
                        string imgName = $"Im{imgId}";
                        _images[imgName] = imgId;

                        using (var msJpeg = new MemoryStream())
                        {
                            var codec = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                            var ep = new EncoderParameters(1);
                            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, _settings.JpegQuality);
                            bmp.Save(msJpeg, codec, ep);
                            byte[] imgData = msJpeg.ToArray();

                            long pos = _fs.Position;
                            _offs[imgId] = pos;
                            byte[] header = Encoding.ASCII.GetBytes($"{imgId} 0 obj\n<< /Type /XObject /Subtype /Image /Width {bmp.Width} /Height {bmp.Height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {imgData.Length} >>\nstream\n");
                            _fs.Write(header, 0, header.Length);
                            _fs.Write(imgData, 0, imgData.Length);
                            byte[] footer = Encoding.ASCII.GetBytes("\nendstream\nendobj\n");
                            _fs.Write(footer, 0, footer.Length);

                            float w = sdb.cxDest;
                            float h = sdb.cyDest;
                            float x = sdb.xDest;
                            float y = sdb.yDest + h; // Инверсия Y через матрицу уже выполнена, берем нижнюю точку

                            _content.AppendLine($"q {w} 0 0 {h} {x} {y} cm /{imgName} Do Q");
                        }
                        DeleteObject(hBmp);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("DrawImage failed: " + ex.Message);
                }
            }

            private static string ColorToPdf(uint c)
            {
                float r = (c & 0xFF) / 255.0f;
                float g = ((c >> 8) & 0xFF) / 255.0f;
                float b = ((c >> 16) & 0xFF) / 255.0f;
                return $"{r.ToString("F3", System.Globalization.CultureInfo.InvariantCulture)} " +
                       $"{g.ToString("F3", System.Globalization.CultureInfo.InvariantCulture)} " +
                       $"{b.ToString("F3", System.Globalization.CultureInfo.InvariantCulture)}";
            }

            // Удалено: private static float FlipY(int y) => -y; // Теперь обрабатывается матрицей cm
        }

        #region zlib (PDF FlateDecode = RFC1950: 0x78 0x9C + raw deflate + adler32)

        private static byte[] Zlib(byte[] raw)
        {
            using (var ms = new MemoryStream())
            {
                using (var ds = new DeflateStream(ms, CompressionMode.Compress, true))
                    ds.Write(raw, 0, raw.Length);

                byte[] def = ms.ToArray();
                byte[] res = new byte[def.Length + 6];
                res[0] = 0x78;
                res[1] = 0x9C;
                Buffer.BlockCopy(def, 0, res, 2, def.Length);

                uint a = Adler32(raw);
                res[res.Length - 4] = (byte)(a >> 24);
                res[res.Length - 3] = (byte)(a >> 16);
                res[res.Length - 2] = (byte)(a >> 8);
                res[res.Length - 1] = (byte)a;
                return res;
            }
        }

        private static uint Adler32(byte[] data)
        {
            uint a = 1, b = 0;
            for (int i = 0; i < data.Length; i++)
            {
                a = (a + data[i]) % 65521;
                b = (b + a) % 65521;
            }
            return (b << 16) | a;
        }

        #endregion

        private static string Inv(double v)
        {
            return v.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}