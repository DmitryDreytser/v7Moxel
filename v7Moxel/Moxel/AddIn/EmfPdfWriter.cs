// EmfPdfWriter.cs
// Сборка PDF из страниц-расширенных метафайлов (EMF), полученных PrintWrapper'ом.
// Самодостаточен: только .NET Framework 4 + System.Drawing, внешних библиотек (Pechkin/wkhtmltopdf) не требует.
//
// Контракт PrintWrapper → EmfPdfWriter:
//   каждая страница — это байты одного EMF, созданного CreateEnhMetaFile(реальный принтер DC, frame=весь лист).
//   Из ENHMETAHEADER.rclFrame (0.01 мм) берётся MediaBox листа, содержимое растеризуется GDI+ на 24bpp
//   (или в JPEG) и встраивается XObject'ом на всю страницу.
//
// Ограничение v1: PDF растровый (текст не выделяется/не ищется). Векторный конвертер EMF→PDF — отдельная задача v2.
// Кодировка страницы: FlateDecode (zlib, без потерь) с автоопределением grayscale (в 2 раза меньше данных).

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
    /// <summary>Настройки растеризации EMF-страниц в PDF.</summary>
    public sealed partial class EmfPdfSettings
    {
        /// <summary>Плотность растеризации, dpi. 300 — печатное качество, 200 — быстрее и легче в ~2 раза.</summary>
        public int Dpi = 150;

        /// <summary>true — JPEG (DCTDecode); false — без потерь (FlateDecode). Для таблиц без потерь обычно лучше и легче.</summary>
        public bool UseJpeg = true;

        /// <summary>Качество JPEG, если UseJpeg.</summary>
        public int JpegQuality = 92;

        /// <summary>true — если страница полностью серая (r==g==b), писать 8bpp DeviceGray вместо 24bpp DeviceRGB.</summary>
        public bool AllowGray = true;
    }

    public static class EmfPdfWriter
    {
        #region GDI P/Invoke (самодостаточно — WinApi.cs не трогаем)

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr SetEnhMetaFileBits(uint nSize, byte[] lpData);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetEnhMetaFileHeader(IntPtr hEmf, uint nSize, out PrintWrapper.ENHMETAHEADER hdr);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetEnhMetaFileBits(IntPtr hEmf, uint nSize, byte[] lpData);

        // ENHMETAHEADER (wingdi.h, x86, 80 байт):
        //  0: iType(4),  4: nSize(4),  8: rclBounds(16), 24: rclFrame(16, 0.01мм),
        // 40: dSignature=0x464D4520, 44: nVersion, 48: nBytes, 52: nRecords, ...
        private const int EMH_RCLFRAME = 24;
        private const uint ENHMETA_SIGNATURE = 0x464D4520;

        #endregion

        /// <summary>
        /// Собирает PDF из страниц EMF. Возвращает true при успехе.
        /// </summary>
        public static bool Save(IList<byte[]> emfPages, string fileName, EmfPdfSettings settings = null)
        {
            if (emfPages == null || emfPages.Count == 0 || string.IsNullOrEmpty(fileName))
                return false;

            if (settings == null)
                settings = new EmfPdfSettings();

            int dpi = settings.Dpi <= 0 ? 300 : settings.Dpi;

            try
            {
                int count = emfPages.Count;

                // Раскладка объектов: 1=Catalog, 2=Pages, для страницы i: (3+3i)=Page, (4+3i)=Contents, (5+3i)=Image
                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, 256 * 1024))
                {
                    long[] offs = new long[3 * count + 3];

                    Action<string> W = s =>
                    {
                        byte[] b = System.Text.Encoding.ASCII.GetBytes(s);
                        fs.Write(b, 0, b.Length);
                    };
                    Action<byte[]> WB = b => fs.Write(b, 0, b.Length);

                    Action<int, Action> Obj = (id, body) =>
                    {
                        offs[id] = fs.Position;
                        W(id + " 0 obj\n");
                        body();
                        W("\nendobj\n");
                    };

                    // --- Заголовок (%%EOF-комментарий с бинарными байтами дляASCII-безопасности) ---
                    W("%PDF-1.4\n");
                    WB(new byte[] { 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A }); // %âãÏÓ

                    Obj(1, () => W("<< /Type /Catalog /Pages 2 0 R >>"));

                    Obj(2, () =>
                    {
                        var kids = new StringBuilder();
                        for (int i = 0; i < count; i++)
                            kids.Append(3 + 3 * i).Append(" 0 R ");
                        W("<< /Type /Pages /Count " + count + " /Kids [ " + kids + "] >>");
                    });

                    for (int i = 0; i < count; i++)
                    {
                        // --- Растеризация страницы ---
                        RenderedPage pg;
                        try
                        {
                            pg = RenderPage(emfPages[i], dpi, settings, Path.ChangeExtension(fileName,$"{i}.wmf"));
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"EmfPdfWriter: страница {i + 1} не растеризована: {ex.Message}");
                            continue;
                        }

                        if (pg == null)
                            continue;

                        double mediaW = Math.Round(pg.MediaWidthPt, 2);
                        double mediaH = Math.Round(pg.MediaHeightPt, 2);
                        if (mediaW <= 0 || mediaH <= 0)
                        {
                            // На всякий случай: A4
                            mediaW = 595.28;
                            mediaH = 841.89;
                        }

                        // --- Page ---
                        int pageId = 3 + 3 * i;
                        int contId = 4 + 3 * i;
                        int imgId = 5 + 3 * i;

                        string content = $"q\n{Inv(pg.MediaWidthPt)} 0 0 {Inv(pg.MediaHeightPt)} 0 0 cm\n/Im{i} Do\nQ\n";
                        byte[] contentBytes = Encoding.ASCII.GetBytes(content);

                        Obj(pageId, () =>
                            W($"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {Inv(mediaW)} {Inv(mediaH)}] " +
                              $"/Resources << /XObject << /Im{i} {imgId} 0 R >> /ProcSet [/PDF /ImageB /ImageC] >> " +
                              $"/Contents {contId} 0 R >>"));

                        Obj(contId, () =>
                        {
                            W($"<< /Length {contentBytes.Length} >>\nstream\n");
                            WB(contentBytes);
                            W("\nendstream");
                        });

                        string filter = pg.Jpeg ? "/DCTDecode" : "/FlateDecode";
                        string cs = pg.Gray ? "/DeviceGray" : "/DeviceRGB";

                        Obj(imgId, () =>
                        {
                            W($"<< /Type /XObject /Subtype /Image /Width {pg.PixelWidth} /Height {pg.PixelHeight} " +
                              $"/ColorSpace {cs} /BitsPerComponent 8 /Filter {filter} /Length {pg.Payload.Length} >>\nstream\n");
                            WB(pg.Payload);
                            W("\nendstream");
                        });
                    }

                    // --- xref / trailer ---
                    long xrefPos = fs.Position;
                    int totalObjects = 3 * count + 2;

                    W($"xref\n0 {totalObjects + 1}\n");
                    W("0000000000 65535 f \n");
                    for (int id = 1; id <= totalObjects; id++)
                        W($"{offs[id]:D10} 00000 n \n");

                    W($"trailer\n<< /Size {totalObjects + 1} /Root 1 0 R >>\n");
                    W($"startxref\n{xrefPos}\n%%EOF");
                }

                return new FileInfo(fileName).Length > 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"EmfPdfWriter.Save: {ex}");
                return false;
            }
        }

        private sealed class RenderedPage
        {
            public byte[] Payload;
            public bool Jpeg;
            public bool Gray;
            public int PixelWidth, PixelHeight;
            public double MediaWidthPt, MediaHeightPt;
        }

        /// <summary>Растеризует один EMF в Bitmap и кодирует во встраиваемый payload.</summary>
        private static RenderedPage RenderPage(byte[] emf, int dpi, EmfPdfSettings settings, string emfFile = null)
        {
            IntPtr hmf = SetEnhMetaFileBits((uint)emf.Length, emf);
            if (hmf == IntPtr.Zero)
                throw new InvalidOperationException("SetEnhMetaFileBits failed");

            Metafile mf = null;
            try
            {
                // rclFrame → размер листа в пунктах (MediaBox) и в пикселях растра
                if (GetEnhMetaFileHeader(hmf, 88, out var hdr) == 0)
                    throw new InvalidOperationException("GetEnhMetaFileHeader failed");

                if (hdr.dSignature != unchecked((int)ENHMETA_SIGNATURE))
                    throw new InvalidOperationException("Не EMF: битая сигнатура");

                double fL = hdr.rclFrame.left,  fT = hdr.rclFrame.top;

                double fR = hdr.rclFrame.right, fB = hdr.rclFrame.bottom;

                double frameX01mm = Math.Max(1, fR - fL); // 0.01 мм
                double frameY01mm = Math.Max(1, fB - fT);

                double mmW = frameX01mm / 100.0;          // 210.0 мм
                double mmH = frameY01mm / 100.0;          // 297.0 мм

                double mediaWpt = mmW * 72.0 / 25.4;      // 595.3 pt  (можно и старый вариант — он был верный)
                double mediaHpt = mmH * 72.0 / 25.4;      // 841.9 pt

                int pxW = (int)Math.Round(mmW / 25.4 * dpi);
                int pxH = (int)Math.Round(mmH / 25.4 * dpi);

                var res = new RenderedPage
                {
                    PixelWidth = pxW,
                    PixelHeight = pxH,
                    MediaWidthPt = mediaWpt,
                    MediaHeightPt = mediaHpt,
                    Jpeg = settings.UseJpeg
                };

                mf = new Metafile(hmf, true); // Dispose метафайла удалит и hmf
                hmf = IntPtr.Zero;
                if(!string.IsNullOrEmpty(emfFile))
                    mf.Save(emfFile, ImageFormat.Wmf);

                using (var bmp = new Bitmap(pxW, pxH, PixelFormat.Format24bppRgb))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.Clear(Color.White);
                        // GDI+ проигрывает записи EMF (векторно) прямо в растр нужного размера
                        g.DrawImage(mf, new Rectangle(0, 0, pxW, pxH));
                    }

                    if (res.Jpeg)
                    {
                        res.Payload = EncodeJpeg(bmp, settings.JpegQuality);
                    }
                    else
                    {
                        bool gray;
                        byte[] raw = ExtractRaw(bmp, settings.AllowGray, out gray);
                        res.Gray = gray;
                        res.Payload = Zlib(raw);
                    }
                }

                return res;
            }
            finally
            {
                if (mf != null)
                    mf.Dispose();
                if (hmf != IntPtr.Zero)
                    DeleteEnhMetaFileSafe(hmf);
            }
        }

        [DllImport("gdi32.dll")]
        private static extern bool DeleteEnhMetaFile(IntPtr hEmf);

        private static void DeleteEnhMetaFileSafe(IntPtr hmf)
        {
            try { DeleteEnhMetaFile(hmf); }
            catch { /* не критично */ }
        }

        /// <summary>
        /// Копирует 24bpp-растр в RGB-массив (без stride-выравнивания), попутно определяя,
        /// что страница полностью серая — тогда отдаёт 8bpp gray-массив.
        /// </summary>
        private static byte[] ExtractRaw(Bitmap bmp, bool allowGray, out bool isGray)
        {
            int w = bmp.Width, h = bmp.Height;
            Rectangle rect = new Rectangle(0, 0, w, h);
            BitmapData bd = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            try
            {
                int stride = bd.Stride;
                byte[] row = new byte[stride];
                byte[] rgb = new byte[w * h * 3];
                byte[] gray = allowGray ? new byte[w * h] : null;
                isGray = allowGray;

                int pos = 0;
                for (int y = 0; y < h; y++)
                {
                    Marshal.Copy(bd.Scan0 + y * stride, row, 0, stride);
                    int src = 0;
                    for (int x = 0; x < w; x++)
                    {
                        byte b = row[src], g2 = row[src + 1], r = row[src + 2];
                        src += 3;
                        rgb[pos] = r;
                        rgb[pos + 1] = g2;
                        rgb[pos + 2] = b;
                        pos += 3;

                        if (isGray)
                        {
                            if (r == g2 && g2 == b)
                                gray[y * w + x] = r;
                            else
                                isGray = false;
                        }
                    }
                }

                return isGray ? gray : rgb;
            }
            finally
            {
                bmp.UnlockBits(bd);
            }
        }

        private static byte[] EncodeJpeg(Bitmap bmp, int quality)
        {
            ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders()
                .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
            using (var ep = new EncoderParameters(1))
            {
                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)Math.Max(1, Math.Min(100, quality)));
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, codec, ep);
                    return ms.ToArray();
                }
            }
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
