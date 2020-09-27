using Pechkin;
using Pechkin.Synchronized;
using System.IO;
using System.Drawing.Printing;
using System.Text;
using System.Windows.Forms;
using System;
using DocumentFormat.OpenXml.Bibliography;
using System.Threading.Tasks;

namespace Moxel
{
    public static class PDFWriter
    {

        private static System.Timers.Timer timer = new System.Timers.Timer(100);
        public static event ConverterProgressor onProgress;

        public static GlobalConfig gc = null;

        public static Margins margins = new Margins(15, 15, 12, 12);
        public static bool landscape = false;
        public static PaperKind paperKind = PaperKind.A4;

        public static async Task<bool> Save(Moxel moxel, string filename, ObjectConfig options = null)
        {
            string HTML = HtmlWriter.RenderToHtml(moxel).ToString();


            if (Converter.PageSettings != null)
            {
                gc = new GlobalConfig();

                margins = new Margins(
                    (int)((int)Converter.PageSettings.Get(PageSettings.OptionType.Left) / 25.4 * 100),
                    (int)((int)Converter.PageSettings.Get(PageSettings.OptionType.Right) / 25.4 * 100),
                    (int)((int)Converter.PageSettings.Get(PageSettings.OptionType.Top) / 25.4 * 100),
                    (int)((int)Converter.PageSettings.Get(PageSettings.OptionType.Bottom) / 25.4 * 100)
                    );

                gc.SetMargins(margins).
                    SetPaperSize((PaperKind)Converter.PageSettings.Get(PageSettings.OptionType.Paper)).
                    SetPaperOrientation((int)Converter.PageSettings.Get(PageSettings.OptionType.Orient) == 2);

                options = new ObjectConfig();

                options.SetZoomFactor(10);

                if ((int)Converter.PageSettings.Get(PageSettings.OptionType.FitToPage) == 1)
                    options.SetIntelligentShrinking(false);
                //else
                //    options.SetZoomFactor((int)Converter.PageSettings.Get(PageSettings.OptionType.Scale) / 100 * 10);

                gc.SetColorMode((int)Converter.PageSettings.Get(PageSettings.OptionType.BlackAndWhite) == 1);

                //options.Header.SetLeftText(Encoding.Unicode.GetString(Encoding.Convert(Encoding.GetEncoding(1251), Encoding.Unicode, Encoding.ASCII.GetBytes(moxel.Header.Text))));
                //options.Header.SetFont(new System.Drawing.Font("Arial Unicode MS", 8));
                //options.Footer.SetLeftText(Encoding.Unicode.GetString(Encoding.Convert(Encoding.GetEncoding(1251), Encoding.Unicode, Encoding.ASCII.GetBytes(moxel.Footer.Text))));
                //options.Footer.SetFont(new System.Drawing.Font("Arial Unicode MS", 8));

            }
            else
            {
                if (gc == null)
                    gc = new GlobalConfig();

                if (options == null)
                {
                    if (moxel.GetWidth(0, moxel.nAllColumnCount) * 0.875 > moxel.GetHeight(0, moxel.nAllRowCount) / 3)
                        landscape = true;
                    else
                        landscape = false;

                    gc.SetMargins(margins).SetPaperSize(paperKind).SetPaperOrientation(landscape);

                    options = new ObjectConfig();
                    options.SetIntelligentShrinking(false);
                    options.SetZoomFactor(10);
                }
            }


            IPechkin pechkin = new SynchronizedPechkin(gc);
            pechkin.ProgressChanged += Pechkin_ProgressChanged;
            timer.Elapsed += Timer_Tick;
            timer.Start();
            byte[] buffer = await Task.Factory.StartNew( ()=>  pechkin.Convert(options, HTML));
            timer.Stop();
            if (buffer.Length > 0)
                File.WriteAllBytes(filename, buffer);

            pechkin.ProgressChanged -= Pechkin_ProgressChanged;
            timer.Elapsed -= Timer_Tick;
            return File.Exists(filename);

        }

        private static void Timer_Tick(object sender, EventArgs e)
        {
            onProgress?.Invoke(lastProgress);
        }

        static int lastProgress = 0;
        private static void Pechkin_ProgressChanged(SimplePechkin converter, int progress, string progressDescription)
        {
            //Application.RaiseIdle(EventArgs.Empty);
            onProgress?.Invoke(progress);
            lastProgress = progress;
        }
    }
}
