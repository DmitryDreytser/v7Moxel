using Moxel;
using PDFjet.NET;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static Moxel.Moxel;

namespace Moxel
{
    public class Paper
    {
        private static float InchHundToPoints(float hOfInches)
        {
            return hOfInches / 100f * 72f;
        }

        public readonly float[] DIMENSIONS;

        public Paper(PaperSize paper, bool landscape)
        {
            DIMENSIONS = landscape ? new[] { InchHundToPoints(paper.Height), InchHundToPoints(paper.Width) } : new[] { InchHundToPoints(paper.Width), InchHundToPoints(paper.Height) };
        }
    }

    internal static class MoxelExtentions
    {
        public static Font GetFont(this DataCell moxelCell, string defFontName, int defFontSize, PDF pdf)
        {
            return moxelCell.FormatCell.GetFont(defFontName, defFontSize, moxelCell.Parent, pdf);
        }

        public static Font GetFont(this CSheetFormat format, string defFontName, int defFontSize, Moxel parent, PDF pdf)
        {
            var name = (format.dwFlags.HasFlag(MoxelCellFlags.FontName)) ? parent.FontList[format.wFontNumber].lfFaceName : defFontName;

            var size = (format.dwFlags.HasFlag(MoxelCellFlags.FontSize)) ? -format.wFontSize / 4 : defFontSize;


            Font f1 = new Font(pdf,
        new FileStream(
                "C:/Users/madda/source/repos/JetPdf/fonts/OpenSans/OpenSans-Regular.ttf.stream",
                FileMode.Open,
                FileAccess.Read),
        Font.STREAM);

            var font = f1;// new Font(pdf, CoreFont.COURIER);
            font.SetSize(size);

            return font;
        }


    }

    public static class PDFWriterManaged
    {
        private static System.Timers.Timer timer = new System.Timers.Timer(100);
        public static event ConverterProgressor onProgress;
        static int lastProgress = 0;
        private static void Timer_Tick(object sender, EventArgs e)
        {
            onProgress?.Invoke(lastProgress);
        }

        static PrinterSettings printerSettings = new PrinterSettings();

        private const double Defcolumnwidth = 40.0d;

        private const float
           Standard1CSymbol =
               105f; //Стандартный символ 1С = 105 твипов. Это ширина символа "0" шрифтом Arial, размера 10


        private static readonly float UnitsPerPixel = WinApi.GetUnitsPerPixel();

        private static double MoxelWidthToPixels(double pt)
        {
            return pt / 8.23d * Standard1CSymbol / UnitsPerPixel;
        }

        private static double MoxelHeightToPixels(double pt)
        {
            return pt / 3;
        }


        private static double PixelsToPoints(double px)
        {
            return px * 0.75;
        }
        public static double GetWidth(Moxel moxel, int columnNumber)
        {
            double columnwidth = -1;

            if (moxel.Columns.ContainsKey(columnNumber))
                if (moxel.Columns[columnNumber].FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                    columnwidth = (double)moxel.Columns[columnNumber].FormatCell.wWidth;

            if (columnwidth == -1)
                if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                    columnwidth = (double)moxel.DefFormat.wWidth;

            if (columnwidth == -1)
                columnwidth = Defcolumnwidth;

            return columnwidth;
        }

        public static bool Save(Moxel moxel, string filename, int PaperWidth = 0, int PaperHeight = 0)
        {
            timer.Elapsed += Timer_Tick;

            if (File.Exists(filename))
                File.Delete(filename);

            var defFontSize = 8;
            var defFontName = "Courier";

            if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                defFontSize = -moxel.DefFormat.wFontSize / 4;

            if (moxel.FontList.Count == 1)
                defFontName = moxel.FontList.First().Value.lfFaceName;

            using (var ms = new MemoryStream())
            {

                using (var bos = new BufferedStream(ms))
                {
                    PDF pdf = new PDF(bos, Compliance.PDF_A_1B);
                    Page page = null;

                    if (Converter.PageSettings != null)
                    {
                        var paperKind = (PaperKind)Converter.PageSettings.Get(PageSettings.OptionType.Paper);

                        var lnadscape = (int)Converter.PageSettings.Get(PageSettings.OptionType.Orient) == 2;

                        var paperSize = printerSettings.PaperSizes.Cast<PaperSize>().FirstOrDefault(x => x.Kind == paperKind);

                        page = new Page(pdf, new Paper(paperSize, lnadscape).DIMENSIONS);

                    }
                    else
                    {
                        page = new Page(pdf, A4.PORTRAIT);
                    }

                    Table table = new Table();

                    List<List<Cell>> tableData = new List<List<Cell>>();
                    for (int rowNumber = 0; rowNumber < moxel.nAllRowCount; rowNumber++)
                    {
                        Interlocked.Exchange(ref lastProgress, rowNumber / moxel.nAllRowCount * 100 );

                        MoxelRow row = null;
                        if (moxel.Rows.ContainsKey(rowNumber))
                            row = moxel.Rows[rowNumber];

                        var pdfRow = new List<Cell>(moxel.nAllColumnCount);

                        double rowHeight = 0;

                        var autoHeight = false;

                        if (row != null)
                        {
                            if (row.Height == 0)
                                autoHeight = true;
                            else
                                rowHeight = row.Height;

                            for (var columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
                            {
                                var moxelCell = row[columnNumber];
                                var text = moxelCell.Text;
                                var cell = new Cell(moxelCell.GetFont(defFontName, defFontSize, pdf), text);

                                cell.SetWidth(GetWidth(moxel, columnNumber));

                                if (!string.IsNullOrEmpty(text))
                                {
                                    
                                }

                                pdfRow.Add(cell);
                            }
                        }
                        else
                        {
                            for (var columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
                            {
                                var cell = new Cell(moxel.DefFormat.GetFont(defFontName, defFontSize, moxel, pdf), string.Empty);
                                pdfRow.Add(cell);
                            }
                        }
                        tableData.Add(pdfRow);
                    }
                    table.SetData(tableData);
                    table.DrawOn(page);

                    pdf.Complete();

                    File.WriteAllBytes(filename, ms.ToArray());
                }

            }
            return File.Exists(filename);

        }

    }
}
