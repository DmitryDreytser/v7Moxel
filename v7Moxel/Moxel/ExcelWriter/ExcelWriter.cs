using System;
using System.Linq;
using static Moxel.Moxel;

using ClosedXML.Excel;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using ClosedXML.Excel.Drawings;
using System.Collections.Generic;
using System.Windows.Forms.VisualStyles;
using System.Diagnostics;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using System.Threading;
using System.Collections.Concurrent;
//using DocumentFormat.OpenXml.Office.Drawing;

namespace Moxel
{
    public static class ExcelWriter
    {
        public static event ConverterProgressor onProgress;
        public static PageSettings PageSettings = null;

        static double PixelHeightToExcel(double pixels)
        {
            return pixels * 0.75d;
        }

        static XLBorderStyleValues GetBorderStyle(BorderStyle moxelBorder)
        {
            switch (moxelBorder)
            {
                case BorderStyle.None:
                    return XLBorderStyleValues.None;
                case BorderStyle.ThinSolid:
                    return XLBorderStyleValues.Thin;
                case BorderStyle.ThinDotted:
                    return XLBorderStyleValues.Dotted;
                case BorderStyle.ThinGrayDotted:
                    return XLBorderStyleValues.Dotted;
                case BorderStyle.ThinDashedShort:
                    return XLBorderStyleValues.Dashed;
                case BorderStyle.ThinDashedLong:
                    return XLBorderStyleValues.DashDotDot;
                case BorderStyle.MediumSolid:
                    return XLBorderStyleValues.Medium;
                case BorderStyle.MediumDashed:
                    return XLBorderStyleValues.MediumDashed;
                case BorderStyle.ThickSolid:
                    return XLBorderStyleValues.Thick;
                case BorderStyle.Double:
                    return XLBorderStyleValues.Double;
                default:
                    return XLBorderStyleValues.Thin;
            }
        }

        static Graphics EmptyGraphics = Graphics.FromHwnd(IntPtr.Zero);
        static float MeasureSymbol(Font font)
        {
            double s1 = EmptyGraphics.MeasureString("0", font).Width;
            double s2 = EmptyGraphics.MeasureString("00", font).Width;
            double y = 2 * s1 - s2;//Ширина интервала между символами
            double x = s1 - 2 * y;//Ширина символа
            double z = x + y; //Ширина занимаемого места символом в составе строки
            return (float)z;
        }

        static float MeasureExcelDefaultSymbol()
        {
            return (float)Math.Floor(EmptyGraphics.MeasureString("0", DefaultExcelFont).Width);
        }

        static Size MeasureString(string text, Font font, int AreaWidth)
        {
            Size Constr = new Size { Width = AreaWidth, Height = 0 };
            return System.Windows.Forms.TextRenderer.MeasureText(text, font, Constr, System.Windows.Forms.TextFormatFlags.WordBreak | System.Windows.Forms.TextFormatFlags.VerticalCenter | System.Windows.Forms.TextFormatFlags.ExpandTabs);
        }

        const float standard1CSymbol = 105f; //Стандартный символ 1С = 105 твипов. Это ширина символа "0" шрифтом Arial, размера 10
        static float standardExcelSymbol = 12;   // Стандартный символ Экселя в пикселях. Шрифт по умолчанию - Calibri, размера 11
        static float UnitsPerPixel = WinApi.GetUnitsPerPixel();

        static Font DefaultExcelFont = null;

        static double MoxelWidthToExcel(double pt)
        {
            double Pixels = MoxelWidthToPixels(pt);

            if (Pixels <= standardExcelSymbol)
                return Pixels / standardExcelSymbol;
            else
                return 1 + (Pixels - standardExcelSymbol) / standard1CSymbol * UnitsPerPixel;
        }

        static double MoxelWidthToPixels(double pt)
        {
            return pt / 8d * standard1CSymbol / UnitsPerPixel;
        }

        static double MoxelHeightToPixels(double pt)
        {
            return pt / 3;
        }

        static double MoxelHeightToExcel(double pt)
        {
            return pt / 4;
        }

        static short PixelToMoxelHeight(int pixels)
        {
            return (short)(pixels * 3);
        }

        const double defcolumnwidth = 40.0d;
        static AutoResetEvent finishEvent = new AutoResetEvent(false);
        public static bool Save(Moxel moxel, string filename)
        {

            if (File.Exists(filename))
                File.Delete(filename);

            var f = File.Open(filename, FileMode.OpenOrCreate);
            f.Close();

            int DefFontSize = 8;
            string defFontName = "Arial";

            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                DefaultExcelFont = new Font(workbook.Style.Font.FontName, (float)workbook.Style.Font.FontSize);
                standardExcelSymbol = MeasureExcelDefaultSymbol();
                
                if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                    DefFontSize = -moxel.DefFormat.wFontSize / 4;

                if (moxel.FontList.Count == 1)
                    defFontName = moxel.FontList.First().Value.lfFaceName;

                workbook.Style.Font.FontName = defFontName;
                workbook.Style.Font.FontSize = DefFontSize;

                var worksheet = workbook.Worksheets.Add("Лист1");               

                List<int> RowsToAutoHeight = new List<int>();

                var locker = new AutoResetEvent(true);
                

                Parallel.For(0, moxel.nAllColumnCount, (int columnNumber) =>
                {
                    double columnwidth = -1;

                    if (moxel.Columns.ContainsKey(columnNumber))
                        if (moxel.Columns[columnNumber].FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                            columnwidth = (double)moxel.Columns[columnNumber].FormatCell.wWidth;

                    if (columnwidth == -1)
                        if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                            columnwidth = (double)moxel.DefFormat.wWidth;

                    if (columnwidth == -1)
                        columnwidth = defcolumnwidth;

                    lock (worksheet)
                        worksheet.Column(columnNumber + 1).Width = MoxelWidthToExcel(columnwidth);
                    
                });

                foreach (CellsUnion union in moxel.Unions.Where(x => x.dwBottom != x.dwTop))
                    worksheet.Range(union.dwTop + 1, union.dwLeft + 1, union.dwBottom + 1, union.dwRight + 1).Merge();


                for (int rowNumber = 0; rowNumber < moxel.nAllRowCount; rowNumber++)
                {
                    foreach (CellsUnion union in moxel.Unions.Where(x => x.dwTop == rowNumber && x.dwTop == x.dwBottom))
                        worksheet.Range(union.dwTop + 1, union.dwLeft + 1, union.dwBottom + 1, union.dwRight + 1).Merge();
                }

                foreach (var RowValue in moxel.Rows)
                {
                    var Row = RowValue.Value;
                    var rowNumber = RowValue.Key;

                    foreach (int columnNumber in Row.Keys)
                    {
                        var moxelCell = Row[columnNumber];
                            if (moxelCell.FormatCell.bControlContent == TextControl.Wrap && !moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                                if (string.IsNullOrEmpty(Row[columnNumber + 1].Text) && !string.IsNullOrEmpty(Row[columnNumber].Text))
                                {
                                    var cn = columnNumber + 1;

                                    while (string.IsNullOrEmpty(Row[cn].Text) && cn < Row.Count &&  !worksheet.Cell(rowNumber + 1, cn + 1).IsMerged())
                                        cn++;

                                    if (cn > columnNumber + 1)
                                        worksheet.Range(rowNumber + 1, columnNumber + 1, rowNumber + 1, cn).Merge();
                                }
                    }
                };

                

                var progress = 0;
                var progressor = 0;
                var count = 0;

                ConcurrentDictionary<int, IXLStyle> styleCache = new ConcurrentDictionary<int, IXLStyle>();

                Parallel.For(0, moxel.nAllRowCount, async (int rowNumber) =>
                //for (int rowNumber = 0; rowNumber < moxel.nAllRowCount; rowNumber++)
                {
                    MoxelRow Row = null;
                    if (moxel.Rows.ContainsKey(rowNumber))
                        Row = moxel.Rows[rowNumber];

                    double rowHeight = 0;

                    bool AutoHeight = false;

                    if (Row != null)
                    {

                        if (Row.Height == 0)
                            AutoHeight = true;
                        else
                            rowHeight = Row.Height;

                        bool MeasureAllROw = false;
                        var totalCells = Row.Keys.Count;
                        var RowFinish = new AutoResetEvent(false);

                        Parallel.ForEach(Row.Keys,  (columnNumber) =>
                        //foreach(var columnNumber in Row.Keys)
                        {
                            var moxelCell = Row[columnNumber];
                            IXLRange cell;

                            lock (worksheet)
                                cell = worksheet.Cell(rowNumber + 1, columnNumber + 1).IsMerged() ?
                                    worksheet.Cell(rowNumber + 1, columnNumber + 1).MergedRange() :
                                    worksheet.Cell(rowNumber + 1, columnNumber + 1).AsRange();

                            {
                                var style = cell.Style;
                                string Text = moxelCell.Text;

                                if (!string.IsNullOrEmpty(Text))
                                {
                                    cell.SetValue<string>(Text.TrimEnd('\r', '\n'));
                                    style.Alignment.WrapText = false;
                                    Char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];

                                    int Dots = Text.ToCharArray().Count(t => t == '.');
                                    int Commas = Text.ToCharArray().Count(t => t == ',');

                                    if (Dots > 0)
                                    {
                                        if (Text.Contains(",") && Dots == 1)
                                        {
                                            string tText = Text.Replace(",", "").Replace('.', separator);
                                            double val = 0;
                                            if (Double.TryParse(tText, out val))
                                            {
                                                cell.Value = val;
                                                cell.DataType = XLDataType.Number;
                                                style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2WithSeparator);
                                            }
                                        }
                                    }
                                    else
                                    {

                                        if (Commas > 0)
                                        {
                                            if (Commas == 1)
                                            {
                                                double val = 0;

                                                if (Double.TryParse(Text.Replace(" ", "").Replace(',', separator), out val))
                                                {
                                                    cell.Value = val;
                                                    cell.DataType = XLDataType.Number;
                                                    style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2WithSeparator);
                                                }
                                            }
                                        }
                                    }

                                        lock (style)
                                        {
                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontName))
                                                style.Font.FontName = moxel.FontList[moxelCell.FormatCell.wFontNumber].lfFaceName;
                                            else
                                                style.Font.FontName = defFontName;


                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                                                style.Font.FontSize = -moxelCell.FormatCell.wFontSize / 4;
                                            else
                                                style.Font.FontSize = DefFontSize;

                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontColor))
                                                style.Font.FontColor = XLColor.FromColor(moxelCell.FormatCell.FontColor);
                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontWeight))
                                                if (moxelCell.FormatCell.bFontBold == clFontWeight.Bold)
                                                    style.Font.SetBold();
                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontItalic))
                                                if (moxelCell.FormatCell.bFontItalic)
                                                    style.Font.SetItalic();

                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontUnderline))
                                                if (moxelCell.FormatCell.bFontUnderline)
                                                    style.Font.SetUnderline();


                                            style.Alignment.WrapText = moxelCell.Text.Contains("\r\n");




                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Control))
                                            {
                                                if (moxelCell.FormatCell.bControlContent == TextControl.Auto)
                                                    if (string.IsNullOrEmpty(Row[columnNumber + 1].Text))
                                                        style.Alignment.WrapText = false;
                                                    else
                                                        style.Alignment.WrapText = true;

                                                if (moxelCell.FormatCell.bControlContent == TextControl.Wrap)
                                                    style.Alignment.WrapText = true;

                                                if (moxelCell.FormatCell.bControlContent == TextControl.Cut)
                                                {
                                                    style.Alignment.WrapText = false;
                                                }
                                            }

                                            style.Alignment.TextRotation = moxelCell.TextOrientation;


                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH))
                                                switch (moxelCell.FormatCell.bHorAlign)
                                                {
                                                    case TextHorzAlign.Left:
                                                        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                                        break;
                                                    case TextHorzAlign.Right:
                                                        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                                                        break;
                                                    case TextHorzAlign.Center:
                                                        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                                        break;
                                                    case TextHorzAlign.Justify:
                                                        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;
                                                        break;
                                                    case TextHorzAlign.BySelection:
                                                        style.Alignment.Horizontal = XLAlignmentHorizontalValues.General;
                                                        break;
                                                    case TextHorzAlign.CenterBySelection:
                                                        {
                                                            worksheet.Range(rowNumber + 1, 1, rowNumber + 1, moxel.nAllColumnCount).Select();
                                                            worksheet.SelectedRanges.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.CenterContinuous;
                                                            MeasureAllROw = true;
                                                            break;
                                                        }
                                                    default:
                                                        break;
                                                }
                                            else

                                                style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV))
                                                switch (moxelCell.FormatCell.bVertAlign)
                                                {
                                                    case TextVertAlign.Bottom:
                                                        style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                                                        break;
                                                    case TextVertAlign.Middle:
                                                        style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                                        break;
                                                    case TextVertAlign.Top:
                                                        style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                                                        break;
                                                    default:
                                                        break;
                                                }
                                            else
                                                style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                                        }
                                    

                                    if (AutoHeight)
                                        using (Font fn = new Font(style.Font.FontName, (float)style.Font.FontSize, style.Font.Bold ? FontStyle.Bold : FontStyle.Regular))
                                        {
                                            double AreaWidth = 0d;

                                            if (MeasureAllROw)
                                                AreaWidth = MoxelWidthToPixels(moxel.GetWidth(0, moxel.nAllColumnCount));
                                            else
                                                AreaWidth = MoxelWidthToPixels(moxel.GetWidth(columnNumber, columnNumber + cell.ColumnCount() - 1)); //+ moxel.GetColumnWidth(columnNumber)

                                            SizeF stringSize;

                                            if (!style.Alignment.WrapText)
                                            {
                                                lock (EmptyGraphics)
                                                    stringSize = EmptyGraphics.MeasureString(Text, fn);
                                            }
                                            else
                                                stringSize = MeasureString(Text, fn, (int)Math.Round(AreaWidth));

                                            int heigth = (int)Math.Ceiling(stringSize.Height / cell.RowCount() / 1.27 * 4);

                                            rowHeight = Math.Max(Math.Max(heigth, 45), rowHeight);
                                            lock (style)
                                            {
                                                if (style.Alignment.Vertical == XLAlignmentVerticalValues.Bottom && style.Alignment.WrapText && (moxelCell.TextOrientation == 0))
                                                    style.Alignment.Vertical = XLAlignmentVerticalValues.Justify;
                                            }
                                        }
                                }

                                    lock (style)
                                    {
                                        if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderBottom))
                                            style.Border.BottomBorder = GetBorderStyle(moxelCell.FormatCell.bBorderBottom);

                                        if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                                            style.Border.LeftBorder = GetBorderStyle(moxelCell.FormatCell.bBorderLeft);

                                        if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
                                            style.Border.TopBorder = GetBorderStyle(moxelCell.FormatCell.bBorderTop);

                                        if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                                            style.Border.RightBorder = GetBorderStyle(moxelCell.FormatCell.bBorderRight);

                                        if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                                            style.Fill.BackgroundColor = XLColor.FromColor(moxelCell.FormatCell.BgColor);

                                        style.Border.BottomBorderColor = XLColor.FromColor(moxelCell.FormatCell.BorderColor);
                                        style.Border.TopBorderColor = XLColor.FromColor(moxelCell.FormatCell.BorderColor);
                                        style.Border.RightBorderColor = XLColor.FromColor(moxelCell.FormatCell.BorderColor);
                                        style.Border.LeftBorderColor = XLColor.FromColor(moxelCell.FormatCell.BorderColor);
                                    }
                            }

                            Interlocked.Decrement(ref totalCells);
                            if (totalCells == 0)
                                RowFinish.Set();
                        });

                        await Task.Run(() => RowFinish.WaitOne());

                        if (AutoHeight)
                            if (rowHeight > 0)
                                    Row.Height = (short)Math.Round(rowHeight, 0);
                                else
                                    rowHeight = 45;
                    }
                    else
                        rowHeight = 45;

                   lock(worksheet)
                        worksheet.Row(rowNumber + 1).Height = MoxelHeightToExcel(rowHeight);

                    progress = count * 100 / moxel.nAllRowCount;
                    if (progressor != progress)
                    {
                        progressor = progress;
                        Task.Run( () => onProgress?.Invoke(progressor));
                    }

                    System.Threading.Interlocked.Increment(ref count);

                    if (count == moxel.nAllRowCount)
                        finishEvent.Set();
                });

                finishEvent.WaitOne();

                foreach (EmbeddedObject obj in moxel.Objects)
                {
                    using (var ms = new MemoryStream())
                    {
                        switch (obj.Picture.dwType)
                        {
                            case ObjectType.Ole:

                            case ObjectType.Picture:
                                {

                                    var zoomfactorY = (double)obj.pObject.Height / (double)obj.ImageArea.Height;
                                    var zoomfactorX = (double)obj.pObject.Width / (double)obj.ImageArea.Width;

                                    ///Странныый косяк с GDI+. Без такого финта выдает неопознанную ошибку
                                    using (var bmp = new Bitmap(obj.pObject))
                                        bmp.Save(ms, ImageFormat.Png);

                                    if (obj.ImageArea.Height < obj.pObject.Height / zoomfactorY)
                                    {
                                        worksheet.Row(obj.Picture.dwRowEnd).Height += (obj.pObject.Height - obj.ImageArea.Height) / 4;
                                        moxel.Rows[obj.Picture.dwRowEnd - 1].Height = (short)(worksheet.Row(obj.Picture.dwRowEnd).Height * 4);
                                    }


                                    var picture = worksheet.AddPicture(ms, XLPictureFormat.Png, $"D{obj.Picture.dwZOrder}");
                                    var topLeftCell = worksheet.Cell(obj.Picture.dwRowStart + 1, obj.Picture.dwColumnStart + 1);
                                    picture.Placement = XLPicturePlacement.Move;
                                    picture.Width = (int)Math.Round(obj.pObject.Width / zoomfactorX);
                                    picture.Height = (int)Math.Round(obj.pObject.Height / zoomfactorY);
                                    picture.MoveTo(topLeftCell, obj.Picture.dwOffsetLeft / 3, obj.Picture.dwOffsetTop / 3);


                                    break;
                                }
                            case ObjectType.Text:
                                {
                                    break;
                                }
                            case ObjectType.Line:
                                {
                                    break;
                                }
                            case ObjectType.Rectangle:
                                {
                                    break;
                                }
                            default:
                                break;

                        }
                    }
                }

                if (!string.IsNullOrEmpty(moxel.Header.Text))
                    worksheet.PageSetup.Header.Left.AddText(moxel.Header.Text.Replace("#D", "&D").Replace("#T", "&T").Replace("#P", "&P").Replace("#Q", "&N"));

                if (!string.IsNullOrEmpty(moxel.Footer.Text))
                    worksheet.PageSetup.Footer.Left.AddText(moxel.Footer.Text.Replace("#D", "&D").Replace("#T", "&T").Replace("#P", "&P").Replace("#Q", "&N"));

                foreach (int br in moxel.HorisontalPageBreaks)
                    worksheet.PageSetup.AddHorizontalPageBreak(br + 1);

                foreach (int br in moxel.VerticalPageBreaks)
                    worksheet.PageSetup.AddVerticalPageBreak(br + 1);

                if (PageSettings != null)
                {

                    worksheet.PageSetup.SetPageOrientation((XLPageOrientation)PageSettings.Get(PageSettings.OptionType.Orient));
                    worksheet.PageSetup.SetPaperSize((XLPaperSize)PageSettings.Get(PageSettings.OptionType.Paper));

                    if ((int)PageSettings.Get(PageSettings.OptionType.FitToPage) == 1)
                        worksheet.PageSetup.FitToPages(1, 0);
                    else
                        worksheet.PageSetup.SetScale((int)PageSettings.Get(PageSettings.OptionType.Scale));

                    worksheet.PageSetup.BlackAndWhite = (int)PageSettings.Get(PageSettings.OptionType.BlackAndWhite) == 1;

                    worksheet.PageSetup.Margins.Bottom = (int)PageSettings.Get(PageSettings.OptionType.Bottom) / 25.4;
                    worksheet.PageSetup.Margins.Left = (int)PageSettings.Get(PageSettings.OptionType.Left) / 25.4;
                    worksheet.PageSetup.Margins.Right = (int)PageSettings.Get(PageSettings.OptionType.Right) / 25.4;
                    worksheet.PageSetup.Margins.Top = (int)PageSettings.Get(PageSettings.OptionType.Top) / 25.4;
                    worksheet.PageSetup.Margins.Footer = 0;
                    worksheet.PageSetup.Margins.Header = 0;
                }
                
                workbook.SaveAs(filename);
            }

            return true;
        }
    }
}
