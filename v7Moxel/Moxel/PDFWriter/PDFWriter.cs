using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pechkin;
using Pechkin.Synchronized;
using System.IO;
using System.Drawing.Printing;

namespace Moxel
{
    public static class PDFWriter
    {
        public static bool Save(Moxel moxel, string filename, GlobalConfig options)
        {
            string HTML = HtmlWriter.RenderToHtml(moxel).ToString();

            IPechkin pechkin = new SynchronizedPechkin(options);
            byte[] buffer = pechkin.Convert(HTML);
            if (buffer.Length > 0)
                File.WriteAllBytes(filename, buffer);

            return File.Exists(filename);
        }

        public static bool Save(Moxel moxel, string filename)
        {
            GlobalConfig gc = new GlobalConfig();
            gc.SetMargins(new Margins(2, 2, 2, 2)).SetPaperSize(PaperKind.A4).SetPaperOrientation(true);

            return Save(moxel, filename, gc);

        }
    }
}
