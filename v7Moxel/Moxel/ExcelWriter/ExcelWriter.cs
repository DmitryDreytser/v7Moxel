using System;
using System.Linq;
using static Moxel.Moxel;

using ClosedXML.Excel;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using ClosedXML.Excel.Drawings;
using System.Collections.Generic;

namespace Moxel
{
    public static class ExcelWriter
    {
        public delegate void ExcelConverterProgressor(int progress);

        public static event ExcelConverterProgressor onProgress;

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
        static float MeasureSymbol( Font font)
        {
           
            double s1 = EmptyGraphics.MeasureString("0", font).Width;
            double s2 = EmptyGraphics.MeasureString("00", font).Width;
            double y = 2 * s1 - s2;//Ширина интервала между символами
            double x = s1 - 2 * y;//Ширина символа
            double z = x + y; //Ширина занимаемого места символом в составе строки
            return (float) z;
        }

        static float MeasureExcelDefaultSymbol()
        {
            using (Font fn = new Font("Calibri", 11f, FontStyle.Regular))
            {
                return MeasureSymbol(fn) * 15f;
            }
        }

        static float standard1CSymbol = 105; //Стандартный символ 1С = 105 твипов. Это ширина символа "0" шрифтом Arial, размера 10
        static float standardExcelSymbol = MeasureExcelDefaultSymbol(); // Стандартный символ Экселя. Шрфт по умолчанию - Calibri, размера 11
        static float SymbolZoomCoefficient = standard1CSymbol / standardExcelSymbol; //Коэффициент преобразования ширины символа для расчета ширины колонки

        static double WidthToExcel(double pt)
        {
            return pt / 8d * SymbolZoomCoefficient + 0.5 / 15 * 8d * SymbolZoomCoefficient;
        }

        static double WidthToPixels(double pt)
        {
            return pt / 8d * 105d / 15;
        }

        static double HeightToPixels(double pt)
        {
            return pt / 3;
        }

        static double HeightToExcel(double pt)
        {
            return pt / 4;
        }

        static short PixelToHeight(int pixels)
        {
            return (short)(pixels * 3);
        }

        public static bool Save(Moxel moxel, string filename)
        {

            if (File.Exists(filename))
                File.Delete(filename);

            var f = File.Open(filename, FileMode.OpenOrCreate);
            f.Close();

            int RowCount = 0;
            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                int DefFontSize = 8;
                string defFontName = "Arial";

                List<int> RowsToAutoHeight = new List<int>();

                if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                    DefFontSize = -moxel.DefFormat.wFontSize / 4;

                if (moxel.FontList.Count == 1)
                    defFontName = moxel.FontList.First().Value.lfFaceName;

                var worksheet = workbook.Worksheets.Add("Лист1"); 
                for (int columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
                {
                    double defcolumnwidth = 40.0d;
                    double columnwidth = -1;

                    if (moxel.Columns.ContainsKey(columnNumber) )
                        if (moxel.Columns[columnNumber].FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                            columnwidth = (double)moxel.Columns[columnNumber].FormatCell.wWidth;

                    if (columnwidth == -1)
                        if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                            columnwidth = (double)moxel.DefFormat.wWidth;

                    if (columnwidth == -1)
                        columnwidth = defcolumnwidth;

                    worksheet.Column(columnNumber + 1).Width = WidthToExcel(columnwidth);
                }

                foreach (CellsUnion union in moxel.Unions)
                {
                    worksheet.Range(union.dwTop + 1, union.dwLeft + 1, union.dwBottom + 1, union.dwRight + 1).Merge();
                }

                for (int rowNumber = 0; rowNumber < moxel.nAllRowCount; rowNumber++)
                {
                    int progress = (rowNumber + 1 ) * 100 / moxel.nAllRowCount;

                    onProgress?.Invoke(progress);

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


                        foreach (int columnNumber in Row.Keys)
                        {
                            IXLRange cell;
                            if (worksheet.Cell(rowNumber + 1, columnNumber + 1).IsMerged())
                                cell = worksheet.Cell(rowNumber + 1, columnNumber + 1).MergedRange();
                            else
                                cell = worksheet.Cell(rowNumber + 1, columnNumber + 1).AsRange();

                            var moxelCell = Row[columnNumber];
                            string Text = moxelCell.Text;

                            if (!string.IsNullOrEmpty(Text))
                            {
                                cell.SetValue<string>(Text.TrimEnd('\r', '\n'));

                                Char separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];

                                int Dots = Text.ToCharArray().Count(t => t == '.');

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
                                            cell.Style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2WithSeparator);
                                        }
                                    }
                                }

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontName))
                                    cell.Style.Font.FontName = moxel.FontList[moxelCell.FormatCell.wFontNumber].lfFaceName;
                                else
                                    cell.Style.Font.FontName = defFontName;

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                                    cell.Style.Font.FontSize = -moxelCell.FormatCell.wFontSize / 4;
                                else
                                    cell.Style.Font.FontSize = DefFontSize;

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontColor))
                                    cell.Style.Font.FontColor = XLColor.FromColor(moxelCell.FormatCell.FontColor);

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontWeight))
                                    if (moxelCell.FormatCell.bFontBold == clFontWeight.Bold)
                                        cell.Style.Font.SetBold();

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontItalic))
                                    if (moxelCell.FormatCell.bFontItalic)
                                        cell.Style.Font.SetItalic();

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontUnderline))
                                    if (moxelCell.FormatCell.bFontUnderline)
                                        cell.Style.Font.SetUnderline();

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Control))
                                {
                                    if (moxelCell.FormatCell.bControlContent == TextControl.Auto || moxelCell.FormatCell.bControlContent == TextControl.Wrap)
                                        cell.Style.Alignment.WrapText = true;
                                    if (moxelCell.FormatCell.bControlContent == TextControl.Cut)
                                    {
                                        cell.Style.Alignment.WrapText = false;
                                    }
                                }

                                cell.Style.Alignment.TextRotation = moxelCell.TextOrientation;

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH))
                                    switch (moxelCell.FormatCell.bHorAlign)
                                    {
                                        case TextHorzAlign.Left:
                                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                            break;
                                        case TextHorzAlign.Right:
                                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                                            break;
                                        case TextHorzAlign.Center:
                                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                            break;
                                        case TextHorzAlign.Justify:
                                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;
                                            break;
                                        case TextHorzAlign.BySelection:
                                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.General;
                                            break;
                                        case TextHorzAlign.CenterBySelection:
                                            {
                                                worksheet.Range(rowNumber + 1, 1, rowNumber + 1, moxel.nAllColumnCount).Select();
                                                worksheet.SelectedRanges.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.CenterContinuous;
                                                break;
                                            }
                                        default:
                                            break;
                                    }
                                else
                                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV))
                                    switch (moxelCell.FormatCell.bVertAlign)
                                    {
                                        case TextVertAlign.Bottom:
                                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                                            break;
                                        case TextVertAlign.Middle:
                                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                            break;
                                        case TextVertAlign.Top:
                                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                                            break;
                                        default:
                                            break;
                                    }
                                else
                                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;

                                if (AutoHeight)
                                    using (Font fn = new Font(cell.Style.Font.FontName, (float)cell.Style.Font.FontSize, cell.Style.Font.Bold ? FontStyle.Bold : FontStyle.Regular))
                                    {
                                        SizeF stringSize = EmptyGraphics.MeasureString(Text, fn).ToSize();
                                        int strings = Text.TrimEnd('\r','\n').Split('\n').Length;
                                        double AreaWidth = WidthToPixels(moxel.GetWidth(columnNumber, columnNumber + cell.ColumnCount() - 1) + moxel.GetColumnWidth(columnNumber));

                                        if (cell.Style.Alignment.WrapText)
                                        {
                                            strings = (int)Math.Max(Math.Ceiling(stringSize.Width / AreaWidth), strings);
                                        }

                                        int heigth = (int)Math.Ceiling(strings * stringSize.Height / cell.RowCount() / 1.27 * 4); //PixelToHeight((int)Math.Round(strings * stringSize.Height / cell.RowCount())); //* SymbolZoomCoefficient

                                        rowHeight = Math.Max(Math.Max(heigth, 45) , rowHeight);
                                    }
                            }

                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderBottom))
                                cell.Style.Border.BottomBorder = GetBorderStyle(moxelCell.FormatCell.bBorderBottom);
                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                                cell.Style.Border.LeftBorder = GetBorderStyle(moxelCell.FormatCell.bBorderLeft);
                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
                                cell.Style.Border.TopBorder = GetBorderStyle(moxelCell.FormatCell.bBorderTop);
                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                                cell.Style.Border.RightBorder = GetBorderStyle(moxelCell.FormatCell.bBorderRight);

                            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                                cell.Style.Fill.BackgroundColor = XLColor.FromColor(moxelCell.FormatCell.BgColor);

                            cell.Style.Border.BottomBorderColor = XLColor.FromColor(moxelCell.FormatCell.BorderColor);
                            cell.Style.Border.TopBorderColor = XLColor.FromColor(moxelCell.FormatCell.BorderColor);
                            cell.Style.Border.RightBorderColor = XLColor.FromColor(moxelCell.FormatCell.BorderColor);
                            cell.Style.Border.LeftBorderColor = XLColor.FromColor(moxelCell.FormatCell.BorderColor);


                        }
                        if (AutoHeight)
                            if (rowHeight > 0)
                                Row.Height = (short)Math.Round(rowHeight, 0);
                            else
                                rowHeight = 45;
                    }
                    else
                        rowHeight = 45;

                    worksheet.Row(rowNumber + 1).Height = HeightToExcel(rowHeight);
                }

                foreach(EmbeddedObject obj in moxel.Objects)
                {
                    using (var ms = new MemoryStream())
                    {
                        int zoomfactor = obj.Picture.dwType == ObjectType.Ole ? zoomfactor = 3: zoomfactor = 1;
                        switch (obj.Picture.dwType)
                        {
                            case ObjectType.Ole:
                                
                            case ObjectType.Picture:
                                {
                                    obj.pObject.Save(ms, ImageFormat.Png);

                                    if(obj.ImageArea.Height < obj.pObject.Height / zoomfactor)
                                    {
                                        worksheet.Row(obj.Picture.dwRowEnd).Height += (obj.pObject.Height - obj.ImageArea.Height) / 4;
                                        moxel.Rows[obj.Picture.dwRowEnd - 1].Height = (short)(worksheet.Row(obj.Picture.dwRowEnd).Height * 4);
                                    }
                                    

                                    var picture = worksheet.AddPicture(ms, XLPictureFormat.Png, $"D{obj.Picture.dwZOrder}");
                                    var topLeftCell = worksheet.Cell(obj.Picture.dwRowStart + 1, obj.Picture.dwColumnStart + 1);
                                    picture.Placement = XLPicturePlacement.Move;
                                    picture.Width = obj.ImageArea.Width;
                                    picture.Height = obj.ImageArea.Height;
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

                workbook.SaveAs(filename);
            }
            
            return true;
        }
    }
}
