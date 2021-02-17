using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Moxel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using static Moxel.Moxel;
//using DocumentFormat.OpenXml.Drawing;

//using System.Windows.Forms;

//using DocumentFormat.OpenXml.Office.Drawing;

namespace v7Moxel.Moxel.ExcelWriter
{
    internal static class MoxelExtentions
    {
        public static Font GetFont(this DataCell moxelCell, string defFontName, int defFontSize)
        {
            var name = (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontName)) ? moxelCell.Parent.FontList[moxelCell.FormatCell.wFontNumber].lfFaceName
            : defFontName;

            var size = (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontSize)) ? -moxelCell.FormatCell.wFontSize / 4 : defFontSize;

            var fontSyle = FontStyle.Regular;

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontWeight))
                if (moxelCell.FormatCell.bFontBold == clFontWeight.Bold)
                    fontSyle |= FontStyle.Bold;

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontItalic))
                if (moxelCell.FormatCell.bFontItalic)
                    fontSyle |= FontStyle.Italic;

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontUnderline))
                if (moxelCell.FormatCell.bFontUnderline)
                    fontSyle |= FontStyle.Underline;

            return new Font(name, size, fontSyle);
        }

        public static void SetHorisonttalAlign(this DataCell moxelCell, MoxelRow row, ExcelStyle style)
        {
            var moxel = moxelCell.Parent;

            var alignHor =
              moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH)
                  ? moxelCell.FormatCell.bHorAlign
                  : (row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH)
                      ? row.FormatCell.bHorAlign
                      : (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.AlignH)
                          ? moxel.DefFormat.bHorAlign
                          : TextHorzAlign.Left));

            switch (alignHor)
            {
                case TextHorzAlign.Left:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    break;
                case TextHorzAlign.Right:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    break;
                case TextHorzAlign.Center:
                    style.HorizontalAlignment =
                      ExcelHorizontalAlignment.Center;
                    break;
                case TextHorzAlign.Justify:
                   style.HorizontalAlignment =
                      ExcelHorizontalAlignment.Justify;
                    break;
            }
        }
            public static void SetVerticalAlign(this DataCell moxelCell, MoxelRow row, ExcelStyle style)
        {
            var moxel = moxelCell.Parent;

            var alignVert =
              moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV)
                  ? moxelCell.FormatCell.bVertAlign
                  : (row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV)
                      ? row.FormatCell.bVertAlign
                      : (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.AlignV)
                          ? moxel.DefFormat.bVertAlign
                          : TextVertAlign.Bottom));

            switch (alignVert)
            {
                case TextVertAlign.Bottom:
                    style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                    break;
                case TextVertAlign.Middle:
                    style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    break;
                case TextVertAlign.Top:
                    style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    break;
            }
        }
        public static void SetBorder(this DataCell moxelCell, Border border)
        {
            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderBottom))
            {
                border.Bottom.Style =
                  ExcelWriter.GetBorderStyle(moxelCell.FormatCell.bBorderBottom);
                if (border.Bottom.Style != ExcelBorderStyle.None &&
                  moxelCell.FormatCell.BorderColor != Color.Black)
                    border.Bottom.Color.SetColor(
                      moxelCell.FormatCell.BorderColor);
            }

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
            {
                border.Left.Style =
                  ExcelWriter.GetBorderStyle(moxelCell.FormatCell.bBorderLeft);
                if (border.Left.Style != ExcelBorderStyle.None &&
                  moxelCell.FormatCell.BorderColor != Color.Black)
                    border.Left.Color.SetColor(moxelCell.FormatCell
                      .BorderColor);
            }

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
            {
                border.Top.Style =
                  ExcelWriter.GetBorderStyle(moxelCell.FormatCell.bBorderTop);
                if (border.Top.Style != ExcelBorderStyle.None &&
                  moxelCell.FormatCell.BorderColor != Color.Black)
                    border.Top.Color.SetColor(moxelCell.FormatCell
                      .BorderColor);
            }

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
            {
                border.Right.Style =
                  ExcelWriter.GetBorderStyle(moxelCell.FormatCell.bBorderRight);
                if (border.Right.Style != ExcelBorderStyle.None &&
                  moxelCell.FormatCell.BorderColor != Color.Black)
                    border.Right.Color.SetColor(moxelCell.FormatCell
                      .BorderColor);
            }
        }
    }

    public static class ExcelWriter
    {
        public static event ConverterProgressor OnProgress;
        public static PageSettings PageSettings = null;

        public static int MaxDegreeOfParallelism = Environment.ProcessorCount / 4 * 3;

        private static double PixelHeightToExcel(double pixels)
        {
            return pixels * 0.75d;
        }

        internal static ExcelBorderStyle GetBorderStyle(global::Moxel.Moxel.BorderStyle moxelBorder)
        {
            return moxelBorder switch
            {
                BorderStyle.None => ExcelBorderStyle.None,
                BorderStyle.ThinSolid => ExcelBorderStyle.Thin,
                BorderStyle.ThinDotted => ExcelBorderStyle.Dotted,
                BorderStyle.ThinGrayDotted => ExcelBorderStyle.Dotted,
                BorderStyle.ThinDashedShort => ExcelBorderStyle.Dashed,
                BorderStyle.ThinDashedLong => ExcelBorderStyle.DashDotDot,
                BorderStyle.MediumSolid => ExcelBorderStyle.Medium,
                BorderStyle.MediumDashed => ExcelBorderStyle.MediumDashed,
                BorderStyle.ThickSolid => ExcelBorderStyle.Thick,
                BorderStyle.Double => ExcelBorderStyle.Double,
                _ => ExcelBorderStyle.Thin
            };
        }

        private static readonly Graphics EmptyGraphics = Graphics.FromHwnd(IntPtr.Zero);

        private static float MeasureSymbol(Font font)
        {
            double s1 = EmptyGraphics.MeasureString("0", font).Width;
            double s2 = EmptyGraphics.MeasureString("00", font).Width;
            var y = 2 * s1 - s2; //Ширина интервала между символами
            var x = s1 - 2 * y; //Ширина символа
            var z = x + y; //Ширина занимаемого места символом в составе строки
            return (float)z;
        }

        private static float MeasureExcelDefaultSymbol()
        {
            return (float)Math.Floor(EmptyGraphics.MeasureString("0", DefaultExcelFont).Width);
        }

        private static Size MeasureString(string text, Font font, int areaWidth)
        {
            var constr = new Size { Width = areaWidth, Height = 0 };
            return System.Windows.Forms.TextRenderer.MeasureText(text, font, constr,
                System.Windows.Forms.TextFormatFlags.WordBreak | System.Windows.Forms.TextFormatFlags.VerticalCenter |
                System.Windows.Forms.TextFormatFlags.ExpandTabs);
        }

        private const float
            Standard1CSymbol =
                105f; //Стандартный символ 1С = 105 твипов. Это ширина символа "0" шрифтом Arial, размера 10

        private static float
            _standardExcelSymbol = 6f; // Стандартный символ Экселя в пикселях. Шрифт по умолчанию - Calibri, размера 11

        private static readonly float UnitsPerPixel = WinApi.GetUnitsPerPixel();

        private static readonly Font DefaultExcelFont = null;

        private static double MoxelWidthToExcel(double pt)
        {
            var pixels = MoxelWidthToPixels(pt);

            if (pixels <= _standardExcelSymbol)
                return pixels / _standardExcelSymbol;
            else
                return 1 + (pixels - _standardExcelSymbol) / Standard1CSymbol * UnitsPerPixel;
            //return Pixels * 72 / 96;
        }

        private static double MoxelWidthToPixels(double pt)
        {
            return pt / 8.23d * Standard1CSymbol / UnitsPerPixel;
        }

        private static double MoxelHeightToPixels(double pt)
        {
            return pt / 3;
        }

        private static double MoxelHeightToExcel(double pt)
        {
            return pt / 4;
        }

        private static short PixelToMoxelHeight(int pixels)
        {
            return (short)(pixels * 3);
        }

        private const double Defcolumnwidth = 40.0d;
        private static readonly AutoResetEvent FinishEvent = new AutoResetEvent(false);
        

        public static bool Save(global::Moxel.Moxel moxel, string filename)
        {
            _standardExcelSymbol = MeasureSymbol(System.Drawing.SystemFonts.DefaultFont);

            int progress;

            if (File.Exists(filename))
                File.Delete(filename);

            var defFontSize = 8;
            var defFontName = "Arial";
            var separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];

            var styleCache = new ConcurrentDictionary<int, ExcelStyle>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var ms = new MemoryStream())
            using (var package = new ExcelPackage(ms))
            {
                if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                    defFontSize = -moxel.DefFormat.wFontSize / 4;

                if (moxel.FontList.Count == 1)
                    defFontName = moxel.FontList.First().Value.lfFaceName;

                using (var worksheet = package.Workbook.Worksheets.Add("Лист1"))
                {
                    for (var columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
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

                        worksheet.Column(columnNumber + 1).Width = MoxelWidthToExcel(columnwidth);
                    }

                    foreach (var union in moxel.Unions)
                        worksheet.Cells[union.dwTop + 1, union.dwLeft + 1, union.dwBottom + 1, union.dwRight + 1]
                            .Merge = true;

                    var count = 0;
                    var progressor = 0;

                    //var defVertAligne = moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.AlignV) ? moxel.DefFormat.bVertAlign : TextVertAlign.Bottom;

                    //var defHorAligne = moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.AlignH) ? moxel.DefFormat.bHorAlign : TextHorzAlign.Left;

                    foreach (var rowValue in moxel.Rows)
                    {
                        var row = rowValue.Value;
                        var rowNumber = rowValue.Key;

                        foreach (var columnNumber in row.Keys)
                        {
                            var moxelCell = row[columnNumber];

                            if (moxelCell.FormatCell.bControlContent != TextControl.Wrap ||
                                moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight)) continue;

                            if (!string.IsNullOrEmpty(row[columnNumber + 1].Text) ||
                                string.IsNullOrEmpty(row[columnNumber].Text)) continue;

                            var cn = columnNumber + 1;

                            while (row[cn].FormatCell.bHorAlign.HasFlag(TextHorzAlign.BySelection) &&
                                   row[cn].FormatCell.bBorderRight == BorderStyle.None)
                                cn++;

                            if (cn > columnNumber + 1)
                                worksheet.Cells[rowNumber + 1, columnNumber + 1, rowNumber + 1, cn].Merge = true;
                        }
                    };


                    Parallel.For(0, moxel.nAllRowCount, new ParallelOptions { MaxDegreeOfParallelism = MaxDegreeOfParallelism }, (int rowNumber) =>
                      {
                          MoxelRow row = null;
                          if (moxel.Rows.ContainsKey(rowNumber))
                              row = moxel.Rows[rowNumber];

                          double rowHeight = 0;

                          var autoHeight = false;

                          if (row != null)
                          {

                              if (row.Height == 0)
                                  autoHeight = true;
                              else
                                  rowHeight = row.Height;

                              var measureAllRow = false;

                              for (var columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
                              {
                                  var moxelCell = row[columnNumber];

                                  //styleCache.TryGetValue(moxelCell.FormatCell.GetHashCode(), out var exlStyle);

                                  var lastCell = columnNumber;

                                  string range = null;
                                  if (worksheet.Cells[rowNumber + 1, columnNumber + 1, rowNumber + 1, lastCell + 1]
                                    .Merge)
                                      range = worksheet.MergedCells[rowNumber + 1, columnNumber + 1];

                                  using (var cell = range == null
                                    ? worksheet.Cells[rowNumber + 1, columnNumber + 1, rowNumber + 1, lastCell + 1]
                                    : worksheet.Cells[range])
                                  {
                                      var text = moxelCell.Text;

                                      if (!string.IsNullOrEmpty(text))
                                      {
                                          cell.Value = text.TrimEnd('\r', '\n');

                                          cell.Style.WrapText = false;


                                          var dots = text.ToCharArray().Count(t => t == '.');
                                          var commas = text.ToCharArray().Count(t => t == ',');

                                          if (dots > 0 || commas > 0)
                                          {
                                              var tText = commas switch
                                              {
                                                  > 0 when dots == 1 => text.Replace(",", "").Replace('.', separator),
                                                  1 => text.Replace(",", "").Replace(',', separator),
                                                  _ => null
                                              };

                                              if (tText != null && double.TryParse(tText, out var val))
                                              {
                                                  cell.Value = val;
                                                  cell.Style.Numberformat.Format = "#,##0.00";
                                              }
                                          }

                                          cell.Style.Font.SetFromFont(moxelCell.GetFont(defFontName, defFontSize));

                                          if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontColor))
                                              cell.Style.Font.Color.SetColor(moxelCell.FormatCell.FontColor);

                                          cell.Style.WrapText = moxelCell.Text.Contains("\r\n");

                                          if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Control))
                                          {
                                              switch (moxelCell.FormatCell.bControlContent)
                                              {
                                                  case TextControl.Auto
                                                      when string.IsNullOrEmpty(row[columnNumber + 1].Text):
                                                      cell.Style.WrapText = false;
                                                      break;
                                                  case TextControl.Auto:
                                                  case TextControl.Wrap:
                                                      cell.Style.WrapText = true;
                                                      break;
                                                  case TextControl.Cut:
                                                      cell.Style.WrapText = false;
                                                      break;
                                                  case TextControl.Fill:
                                                      break;
                                                  case TextControl.Red:
                                                      break;
                                                  case TextControl.FillAndRed:
                                                      break;
                                                  default:
                                                      throw new ArgumentOutOfRangeException();
                                              }
                                          }

                                          cell.Style.TextRotation = moxelCell.TextOrientation;

                                          moxelCell.SetHorisonttalAlign(row, cell.Style);
                                          moxelCell.SetVerticalAlign(row, cell.Style);

                                          if (autoHeight)
                                              using (var fn = new Font(cell.Style.Font.Name, cell.Style.Font.Size,
                                                cell.Style.Font.Bold ? FontStyle.Bold : FontStyle.Regular))
                                              {
                                                  var areaWidth = 0d;

                                                  if (measureAllRow)
                                                      areaWidth = MoxelWidthToPixels(moxel.GetWidth(0,
                                                        moxel.nAllColumnCount));
                                                  else
                                                      areaWidth = MoxelWidthToPixels(moxel.GetWidth(columnNumber,
                                                        columnNumber + cell.Columns));

                                                  SizeF stringSize;

                                                  if (!cell.Style.WrapText)
                                                  {
                                                      using (var emptyGraphics = Graphics.FromHwnd(IntPtr.Zero))
                                                          stringSize = emptyGraphics.MeasureString(text, fn);
                                                  }
                                                  else
                                                      stringSize = MeasureString(text, fn,
                                                        (int)Math.Round(areaWidth));

                                                  var heigth =
                                                    (int)Math.Ceiling(stringSize.Height / cell.Rows / 1.27 * 4);

                                                  rowHeight = Math.Max(Math.Max(heigth, 45), rowHeight);
                                                  if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV))
                                                      if (cell.Style.VerticalAlignment ==
                                                        ExcelVerticalAlignment.Bottom && cell.Style.WrapText &&
                                                        (moxelCell.TextOrientation == 0))
                                                          cell.Style.VerticalAlignment =
                                                            ExcelVerticalAlignment.Justify;
                                              }
                                      }

                                      try
                                      {
                                          moxelCell.SetBorder(cell.Style.Border);

                                          if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background) &&
                                            moxelCell.FormatCell.BorderColor != Color.White)
                                              {
                                                  cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                  cell.Style.Fill.BackgroundColor.SetColor(moxelCell.FormatCell.BgColor);
                                              }
                                      }
                                      catch (Exception ex)
                                      {

                                          var e = ex;
                                      }


                                  }
                              };

                              if (autoHeight)
                                  if (rowHeight > 0)
                                      row.Height = (short)Math.Round(rowHeight, 0);
                                  else
                                      rowHeight = 45;
                          }
                          else
                              rowHeight = 45;

                          worksheet.Row(rowNumber + 1).Height = MoxelHeightToExcel(rowHeight);

                          progress = (count * 100 / moxel.nAllRowCount);
                          if (progress - progressor > 3)
                          {
                              progressor = progress;
                              OnProgress?.Invoke(progressor);
                          }

                          System.Threading.Interlocked.Increment(ref count);

                          if (count == moxel.nAllRowCount)
                              FinishEvent.Set();
                      });

                    FinishEvent.WaitOne();

                    #region Добавим картинки

                    foreach (var obj in moxel.Objects)
                    {

                        switch (obj.Picture.dwType)
                        {
                            case ObjectType.Ole:

                            case ObjectType.Picture:
                                {
                                    using (var pms = new MemoryStream())
                                    {
                                        //Странныый косяк с GDI+. Без такого финта выдает неопознанную ошибку
                                        using (Bitmap bmp = obj.pObject)
                                        {
                                            bmp.Save(pms, ImageFormat.Png);

                                            var zoomfactorY = (double)bmp.Height / (double)obj.ImageArea.Height;
                                            var zoomfactorX = (double)bmp.Width / (double)obj.ImageArea.Width;

                                            using (var picture = worksheet.Drawings.AddPicture($"D{obj.Picture.dwZOrder}", pms, OfficeOpenXml.Drawing.ePictureType.Png))
                                            {
                                                picture.SetPosition(obj.Picture.dwRowStart, obj.Picture.dwOffsetTop / 3, obj.Picture.dwColumnStart, obj.Picture.dwOffsetLeft / 3);
                                                picture.SetSize((int)Math.Round(bmp.Width / zoomfactorX), (int)Math.Round(bmp.Height / zoomfactorY));
                                            }
                                        }
                                    }
                                    break;
                                }
                            case ObjectType.Text:
                                {

                                    using (var textBox = worksheet.Drawings.AddShape($"D{obj.Picture.dwZOrder}", eShapeStyle.Rect))
                                    {
                                        var row = moxel.Rows[obj.Picture.dwRowStart];

                                        var text = textBox.RichText.Add(obj.Text);
                                        text.SetFromFont(obj.GetFont(defFontName, defFontSize));
                                        
                                        if (obj.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontColor))
                                            textBox.Font.Color = obj.FormatCell.FontColor;
                                        else
                                            textBox.Font.Color = Color.Black;

                                        var area = obj.ImageArea;
                                        
                                        // obj.SetHorisonttalAlign(row, textBox.Style);
                                        // obj.SetVerticalAlign(row, textBox.Style);
                                        
                                        textBox.SetSize(area.Width, area.Height);

                                        if (obj.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                                        {                                            
                                            textBox.Fill.Color = obj.FormatCell.BgColor;
                                        }

                                        textBox.SetPosition(obj.Picture.dwRowStart, obj.Picture.dwOffsetTop / 3, obj.Picture.dwColumnStart, obj.Picture.dwOffsetLeft / 3);
                                    }
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
                            case ObjectType.None:
                                break;
                            default:
                                break;


                        }
                    }

                    #endregion
                    
                    #region Колонтитулы
                    if (!string.IsNullOrEmpty(moxel.Header.Text))
                        worksheet.HeaderFooter.EvenHeader.LeftAlignedText = moxel.Header.Text.Replace("#D", "&D").Replace("#T", "&T")
                            .Replace("#P", "&P").Replace("#Q", "&N");

                    if (!string.IsNullOrEmpty(moxel.Footer.Text))
                        worksheet.HeaderFooter.EvenFooter.LeftAlignedText = moxel.Footer.Text.Replace("#D", "&D").Replace("#T", "&T")
                            .Replace("#P", "&P").Replace("#Q", "&N");
                    #endregion
                    #region Разрывы страниц
                    
                    foreach (int br in moxel.HorisontalPageBreaks)
                        worksheet.Row(br + 1).PageBreak = true;

                    foreach (int br in moxel.VerticalPageBreaks)
                        worksheet.Column(br + 1).PageBreak = true;
                    
                    #endregion

                    #region Параметры страницы

                    if (PageSettings != null)
                    {

                        worksheet.PrinterSettings.Orientation = (eOrientation)((int)PageSettings.Get(PageSettings.OptionType.Orient) - 1);
                        worksheet.PrinterSettings.PaperSize  = (ePaperSize) PageSettings.Get(PageSettings.OptionType.Paper);

                        if ((int) PageSettings.Get(PageSettings.OptionType.FitToPage) == 1)
                            worksheet.PrinterSettings.FitToPage = true;
                        else
                            worksheet.PrinterSettings.Scale = (int) PageSettings.Get(PageSettings.OptionType.Scale);

                        worksheet.PrinterSettings.BlackAndWhite = (int) PageSettings.Get(PageSettings.OptionType.BlackAndWhite) == 1;
                        
                        

                        worksheet.PrinterSettings.BottomMargin = ((int) PageSettings.Get(PageSettings.OptionType.Bottom)) / (decimal) 25.4;
                        worksheet.PrinterSettings.LeftMargin = ((int) PageSettings.Get(PageSettings.OptionType.Left)) / (decimal) 25.4;
                        worksheet.PrinterSettings.RightMargin = ((int) PageSettings.Get(PageSettings.OptionType.Right)) / (decimal) 25.4;
                        worksheet.PrinterSettings.TopMargin = ((int) PageSettings.Get(PageSettings.OptionType.Top)) / (decimal) 25.4;

                        worksheet.PrinterSettings.FooterMargin =
                            ((int) PageSettings.Get(PageSettings.OptionType.Footer)) / (decimal) 25.4;
                        worksheet.PrinterSettings.HeaderMargin = 
                            ((int) PageSettings.Get(PageSettings.OptionType.Header)) / (decimal) 25.4;

                        if((int)PageSettings.Get(PageSettings.OptionType.RepeatRowFrom) != 0)
                            worksheet.PrinterSettings.RepeatRows =
                                new ExcelAddress((int) PageSettings.Get(PageSettings.OptionType.RepeatRowFrom) + 1,
                                    0,
                                    (int) PageSettings.Get(PageSettings.OptionType.RepeatRowTo) + 1,
                                    0);
                        
                        if((int)PageSettings.Get(PageSettings.OptionType.RepeatColFrom) != 0)
                            worksheet.PrinterSettings.RepeatColumns =
                                new ExcelAddress(0,
                                    (int) PageSettings.Get(PageSettings.OptionType.RepeatColFrom) + 1,
                                    0,
                                    (int) PageSettings.Get(PageSettings.OptionType.RepeatColTo) + 1);
                        
                    }
                    #endregion
                    
                    package.Save();
                }

                File.WriteAllBytes(filename, ms.ToArray());
            }

            return File.Exists(filename);
        }
    }
}
