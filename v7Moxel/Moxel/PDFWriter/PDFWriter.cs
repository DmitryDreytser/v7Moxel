using Pechkin;
using Pechkin.Synchronized;
using System.IO;
using System.Drawing.Printing;

namespace Moxel
{
    public static class PDFWriter
    {
        public static GlobalConfig gc = null;

        public static Margins margins = new Margins(2, 0, 2, 2);
        public static bool? landscape = null;
        public static PaperKind paperKind = PaperKind.A4;

        public static bool Save(Moxel moxel, string filename, GlobalConfig options = null)
        {
            string HTML = HtmlWriter.RenderToHtml(moxel).ToString();

            if (landscape == null)
            {
                if (moxel.GetWidth(0, moxel.nAllColumnCount) > moxel.GetHeight(0, moxel.nAllRowCount))
                    landscape = true;
                else
                    landscape = false;
            }

            if (gc == null)
                gc = new GlobalConfig();

            if(options == null)
                gc.SetMargins(margins).SetPaperSize(paperKind).SetPaperOrientation((bool)landscape);
            else
                gc = options;

            IPechkin pechkin = new SynchronizedPechkin(gc);
            byte[] buffer = pechkin.Convert(HTML);
            if (buffer.Length > 0)
                File.WriteAllBytes(filename, buffer);

            return File.Exists(filename);
        }

    }
}
