using System;
using System.Linq;
using static Moxel.Moxel;

using ClosedXML.Excel;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using ClosedXML.Excel.Drawings;

namespace Moxel
{
    public static class ExcelWriter
    {


        static double PixelHeightToExcel(double pixels)
        {
            return pixels * 0.75d;
        }

        static double PixelWidthToExcel(double pixels)
        {
            return (pixels - 12 + 5) / 7d + 1;
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

        public static bool Save(Moxel moxel, string filename, int formatVersion = 7)
        {
            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                int DefFontSize = 8;
                string defFontName = "Arial";


                if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                    DefFontSize = -moxel.DefFormat.wFontSize / 4;

                if (moxel.FontList.Count == 1)
                    defFontName = moxel.FontList.First().Value.lfFaceName;

                var worksheet = workbook.Worksheets.Add("Лист1");
                for (int columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
                {
                    double columnwidth = 40.0d;

                    if (moxel.Columns.ContainsKey(columnNumber))
                        columnwidth = (double)moxel.Columns[columnNumber].FormatCell.wWidth;
                    else
                        if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                        columnwidth = (double)moxel.DefFormat.wWidth;

                    worksheet.Column(columnNumber + 1).Width = columnwidth * 0.107;
                }

                foreach (CellsUnion union in moxel.Unions)
                {
                    worksheet.Range(union.dwTop + 1, union.dwLeft + 1, union.dwBottom + 1, union.dwRight + 1).Merge();
                }

                for (int rowNumber = 0; rowNumber < moxel.nAllRowCount; rowNumber++)
                {
                    MoxelRow Row = null;
                    if (moxel.Rows.ContainsKey(rowNumber))
                        Row = moxel.Rows[rowNumber];

                    double rowHeight = 45;

                    if (Row != null)
                    {
                        foreach (int columnNumber in Row.Keys)
                        {
                            if(Row.Height > 0 )
                                rowHeight = Row.Height;

                            IXLRange cell;
                            if (worksheet.Cell(rowNumber + 1, columnNumber + 1).IsMerged())
                                cell = worksheet.Cell(rowNumber + 1, columnNumber + 1).MergedRange();
                            else
                                cell = worksheet.Cell(rowNumber + 1, columnNumber + 1).AsRange();

                            var moxelCell = Row[columnNumber];
                            string Text = moxelCell.Text;

                            if (!string.IsNullOrWhiteSpace(Text))
                            {
                                cell.Value = Text;

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
                                        cell.Style.Alignment.WrapText = false;
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
                                        default:
                                            break; 
                                    }
                                else
                                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                                if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV))
                                    switch (moxelCell.FormatCell.bVertAlign)
                                    {
                                        case  TextVertAlign.Bottom:
                                            cell.Style.Alignment.Vertical =  XLAlignmentVerticalValues.Bottom;
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

                                


                                if(Row.Height == 0)
                                {
                                    if (!cell.Style.Alignment.WrapText)
                                        Text = Text.Replace(" ", "_");

                                    using (Font fn = new Font(cell.Style.Font.FontName, (float)cell.Style.Font.FontSize, cell.Style.Font.Bold ? FontStyle.Bold : FontStyle.Regular))
                                    {
                                        Size Constr = new Size { Width = (int)Math.Round((moxel.GetWidth(columnNumber, columnNumber + cell.ColumnCount()-1) + moxel.GetColumnWidth(columnNumber)) * 0.875), Height = 0 };
                                        Size textsize = System.Windows.Forms.TextRenderer.MeasureText(Text.TrimStart(' '), fn, Constr, System.Windows.Forms.TextFormatFlags.WordBreak | System.Windows.Forms.TextFormatFlags.NoClipping);
                                        textsize.Height /= cell.RowCount();
                                        rowHeight = Math.Max(Math.Max(textsize.Height * 1.01, 15) * 3, rowHeight);
                                    }
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
                    }
                    if (Row.Height == 0)
                        Row.Height = (short)Math.Round(rowHeight + 3, 0) ;

                    worksheet.Row(rowNumber + 1).Height = rowHeight / 3.787;
                }

                foreach(EmbeddedObject obj in moxel.Objects)
                {
                    using (var ms = new MemoryStream())
                    {
                        obj.pObject.Save(ms, ImageFormat.Png);
                        var picture = worksheet.AddPicture(ms, XLPictureFormat.Png ,$"D{obj.Picture.dwZOrder}");
                        var topLeftCell = worksheet.Cell(obj.Picture.dwRowStart + 1, obj.Picture.dwColumnStart + 1);
                        picture.Placement = XLPicturePlacement.Move;
                        picture.Width = obj.AbsoluteImageArea.Width;
                        picture.Height = obj.AbsoluteImageArea.Height;

                        picture.MoveTo(topLeftCell, (int) Math.Round(obj.Picture.dwOffsetLeft / 2.8), (int)Math.Round(obj.Picture.dwOffsetTop / 2.8));
                      }
                }

                workbook.SaveAs(filename);
            }
            return true;
        }
    }
}
