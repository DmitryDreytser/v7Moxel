using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronPdf;
using System.IO;

namespace Moxel
{
    public static class PDFWriter
    {

        public static bool Save(Moxel moxel, string filename, PdfPrintOptions options)
        {
            HtmlToPdf Renderer = new HtmlToPdf(options);
            string HTML = HtmlWriter.RenderToHtml(moxel).ToString();
            Renderer.RenderHtmlAsPdf(HTML).SaveAs(filename);

            return File.Exists(filename);
        }

        public static bool Save(Moxel moxel, string filename)
        {
            return Save(moxel, filename, new PdfPrintOptions
            {
                CreatePdfFormsFromHtml = true,
                CssMediaType = PdfPrintOptions.PdfCssMediaType.Print,
                FitToPaperWidth = true,
                PaperSize = PdfPrintOptions.PdfPaperSize.A4,
                PaperOrientation = PdfPrintOptions.PdfPaperOrientation.Landscape,
                MarginLeft = 2,
                MarginBottom = 2,
                MarginRight = 2,
                MarginTop = 2,
                Zoom = 70
            }
            );
        }
    }
}
