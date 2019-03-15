using Pechkin;
using Pechkin.Synchronized;
using System.IO;
using System.Drawing.Printing;

namespace Moxel
{
    public static class PDFWriter
    {
        public static GlobalConfig gc = null;

        public static Margins margins = new Margins(15, 15, 12, 12);
        public static bool landscape = false;
        public static PaperKind paperKind = PaperKind.A4;

        public static bool Save(Moxel moxel, string filename, ObjectConfig options = null)
        {
            string HTML = HtmlWriter.RenderToHtml(moxel).ToString();

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


            IPechkin pechkin = new SynchronizedPechkin(gc);
            byte[] buffer = pechkin.Convert(options, HTML);
            if (buffer.Length > 0)
                File.WriteAllBytes(filename, buffer);

            return File.Exists(filename);
        }

    }
}
