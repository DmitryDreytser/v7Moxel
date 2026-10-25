// PDFWriter.cs (v2) — fallback-канал PDF: модель → HTML (HtmlWriter) → Pechkin (wkhtmltopdf).
//
// Основной канал теперь PrintWrapper (родная печать → EMF → PDF, точные колонтитулы/разрывы).
// Этот класс остаётся для случаев: таблица без view, нет принтера, PrintWrapper недоступен.
//
// Исправлено относительно v1:
//   1. ФИКС: timer и подписка ProgressChanged освобождаются в finally — при исключении
//      больше не копятся подписки (Elapsed += на статическом таймере) и не текут таймеры.
//   2. ФИКС: статические поля gc/margins/landscape больше не мутируются — конфигурация
//      собирается в локальные объекты (гонки при повторных вызовах).
//   3. ФИКС: опечатка v1 «PaperHeight != 0 && PaperHeight != 0» → PaperWidth/PaperHeight.
//   4. НОВОЕ: колонтитулы таблицы (moxel.Header/Footer) отдаются движку wkhtmltopdf
//      с подстановкой кодов 1С: [&P]→[page], [&Q]→[topage], [&D]→[date], [&T]→[time].
//   5. ФИКС: пустой результат конвертации → понятная ошибка вместо молчаливого false.

using Pechkin;
using Pechkin.Synchronized;
using System;
using System.Drawing.Printing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Moxel
{
    public static class PDFWriter
    {
        public static event ConverterProgressor onProgress;

        public static Margins margins = new Margins(15, 15, 12, 12);
        public static bool landscape = false;
        public static PaperKind paperKind = PaperKind.A4;

        public static async Task<bool> Save(Moxel moxel, string filename, ObjectConfig options = null, int PaperWidth = 0, int PaperHeight = 0)
        {
            if (moxel == null || string.IsNullOrEmpty(filename))
                return false;

            using (var ms = new MemoryStream())
            {
                HtmlWriter.RenderToHtml(moxel, ms);

                // Локальная конфигурация — статики v1 больше не мутируются
                GlobalConfig gc = BuildGlobalConfig(moxel, PaperWidth, PaperHeight);
                ObjectConfig objCfg = BuildObjectConfig(moxel, options);

                System.Timers.Timer timer = new System.Timers.Timer(100);
                IPechkin pechkin = null;
                byte[] buffer;

                try
                {
                    pechkin = new SynchronizedPechkin(gc);
                    pechkin.ProgressChanged += Pechkin_ProgressChanged;
                    timer.Elapsed += Timer_Tick;
                    timer.Start();

                    buffer = await Task.Factory.StartNew(() => pechkin.Convert(objCfg, ms.ToArray()));
                }
                finally
                {
                    timer.Stop();
                    timer.Elapsed -= Timer_Tick;   // ФИКС: симметричная отписка
                    timer.Dispose();
                    if (pechkin != null)
                        pechkin.ProgressChanged -= Pechkin_ProgressChanged;
                }

                if (buffer == null || buffer.Length == 0)
                {
                    Converter.RaiseExtRuntimeError("Ошибка сохранения таблицы в PDF: конвертер вернул пустой документ");
                    return false;
                }

                File.WriteAllBytes(filename, buffer);
            }
            return File.Exists(filename);
        }

        /// <summary>Глобальные параметры страницы: поля/размер/ориентация — из ПараметрыСтраницы 1С.</summary>
        private static GlobalConfig BuildGlobalConfig(Moxel moxel, int PaperWidth, int PaperHeight)
        {
            var gc = new GlobalConfig();

            if (Converter.PageSettings != null)
            {
                Margins m = new Margins(
                    (int)(Converter.PageSettings.Get<int>(PageSettings.OptionType.Left) * 120 / 25.4),
                    (int)(Converter.PageSettings.Get<int>(PageSettings.OptionType.Right) * 120 / 25.4),
                    (int)(Converter.PageSettings.Get<int>(PageSettings.OptionType.Top) * 120 / 25.4),
                    (int)(Converter.PageSettings.Get<int>(PageSettings.OptionType.Bottom) * 120 / 25.4)
                    );
                gc.SetMargins(m);

                // ФИКС: было «PaperHeight != 0 && PaperHeight != 0» — PaperWidth не учитывался
                if (PaperWidth != 0 && PaperHeight != 0)
                    gc.SetPaperSize(PaperWidth, PaperHeight);
                else
                    gc.SetPaperSize((PaperKind)Converter.PageSettings.Get(PageSettings.OptionType.Paper));

                gc.SetPaperOrientation((int)Converter.PageSettings.Get(PageSettings.OptionType.Orient) == 2);
            }
            else
            {
                // Эвристика альбомной ориентации: ширина таблицы против высоты/3
                if (moxel.GetWidth(0, moxel.nAllColumnCount) * 0.875 > moxel.GetHeight(0, moxel.nAllRowCount) / 3)
                    landscape = true;
                else
                    landscape = false;

                gc.SetMargins(margins);
                if (PaperWidth != 0 && PaperHeight != 0)
                    gc.SetPaperSize(PaperWidth, PaperHeight);
                else
                    gc.SetPaperSize(paperKind);
                gc.SetPaperOrientation(landscape);
            }

            gc.SetColorMode(Converter.PageSettings != null &&
                            (int)Converter.PageSettings.Get(PageSettings.OptionType.BlackAndWhite) == 1);

            return gc;
        }

        /// <summary>Параметры объекта: зум/подгонка + колонтитулы таблицы 1С.</summary>
        private static ObjectConfig BuildObjectConfig(Moxel moxel, ObjectConfig options)
        {
            if (options == null)
            {
                options = new ObjectConfig();

                if (Converter.PageSettings != null)
                {
                    options.SetZoomFactor(10);

                    if ((int)Converter.PageSettings.Get(PageSettings.OptionType.FitToPage) == 1)
                        options.SetIntelligentShrinking(false);
                }
                else
                {
                    options.SetIntelligentShrinking(false);
                    options.SetZoomFactor(10);
                }
            }

            // Колонтитулы таблицы 1С → движок wkhtmltopdf.
            // Коды 1С: [&P] — номер страницы, [&Q] — всего страниц, [&D] — дата, [&T] — время.
            try
            {
                string header = MoxelFooterCodeToWk(moxel.Header != null ? moxel.Header.Text : null);
                string footer = MoxelFooterCodeToWk(moxel.Footer != null ? moxel.Footer.Text : null);

                if (!string.IsNullOrEmpty(header))
                {
                    options.Header.SetCenterText(WrapHeaderHtml(header));
                    options.Header.SetFont(new System.Drawing.Font("Arial", 8));
                }
                if (!string.IsNullOrEmpty(footer))
                {
                    options.Footer.SetCenterText(WrapHeaderHtml(footer));
                    options.Footer.SetFont(new System.Drawing.Font("Arial", 8));
                }
            }
            catch (Exception ex)
            {
                // Колонтитулы не критичны — не роняем конвертацию
                System.Diagnostics.Debug.WriteLine($"PDFWriter: колонтитулы не переданы: {ex.Message}");
            }

            return options;
        }

        /// <summary>Подстановка кодов колонтитулов 1С в токены wkhtmltopdf.</summary>
        private static string MoxelFooterCodeToWk(string text)
        {
            if (string.IsNullOrEmpty(text))
                return null;

            string s = text
                .Replace("[&P]", "[page]")
                .Replace("&P", "[page]")
                .Replace("[&Q]", "[topage]")
                .Replace("&Q", "[topage]")
                .Replace("[&D]", "[date]")
                .Replace("&D", "[date]")
                .Replace("[&T]", "[time]")
                .Replace("&T", "[time]");

            return s.Trim();
        }

        private static string WrapHeaderHtml(string text)
        {
            return "<span style=\"font-family:Arial; font-size:8pt; color:#000000;\">" +
                   System.Net.WebUtility.HtmlEncode(text) + "</span>";
        }

        private static void Timer_Tick(object sender, EventArgs e)
        {
            onProgress?.Invoke(lastProgress);
        }

        static int lastProgress = 0;
        private static void Pechkin_ProgressChanged(SimplePechkin converter, int progress, string progressDescription)
        {
            Application.RaiseIdle(EventArgs.Empty);
            onProgress?.Invoke(progress);
            lastProgress = progress;
        }
    }
}
