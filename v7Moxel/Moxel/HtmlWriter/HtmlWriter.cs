using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using static Moxel.Moxel;

namespace Moxel
{
    /// <summary>
    /// Экспорт документа Moxel в HTML.
    /// Модель документа (Moxel/MoxelRow/DataCell) при экспорте не изменяется.
    /// </summary>
    public class HtmlWriter
    {
        /// <summary>Прогресс экспорта, 0..100.</summary>
        public static event ConverterProgressor onProgress;

        // --- Константы масштабирования (бывшие магические числа) ---
        private const double ColumnWidthToPixels = 0.875; // ширина колонки -> px (в разметке)
        private const double MeasureWidthScale = 0.873; // ширина колонки -> px (при измерении текста)
        private const double QuarterPointToPoints = 0.25;  // четверть пункта -> пункты
        private const int QuarterPointsPerPixel = 3;     // 1 px = 0.75 pt = 3 четверти пункта
        private const int DefaultRowHeightQP = 45;    // высота строки по умолчанию, четверти пункта (11.25 pt)
        private const int DefaultColumnWidthPx = 35;
        private const float DefaultFontSizePt = 8.0f;

        private static readonly List<CellsUnion> NoUnions = new List<CellsUnion>();
        private static int _patternCounter; // счётчик для уникальных id SVG-паттернов

        #region Стиль

        public class CSSstyle : Dictionary<string, string>
        {
            public override string ToString()
            {
                if (Count == 0)
                    return string.Empty;
                return $" style=\"{string.Join("; ", this.Select(kv => $"{kv.Key}: {kv.Value}"))}\"";
            }

            /// <summary>Индексатор словаря и так добавляет ключ — Set оставлен для совместимости.</summary>
            public void Set(string key, string value) => this[key] = value;
        }

        #endregion

        #region Утилиты

        private static string Inv(double value) => value.ToString(CultureInfo.InvariantCulture);
        private static string Rgb(Color color) => $"rgb({color.R},{color.G},{color.B})";
        private static string CssFontName(string name) => "'" + name.Replace("'", "\\'") + "'";

        #endregion

        #region Заполнение стилей

        public static string FillTextStyle(DataCell FormatCell, ref CSSstyle CellStyle)
        {
            string FontFamily = string.Empty;
            float FontSize = 0;
            return FillTextStyle(FormatCell, ref CellStyle, ref FontFamily, ref FontSize);
        }

        public static string FillTextStyle(DataCell FormatCell, ref CSSstyle CellStyle, ref string FontFamily, ref float FontSize)
        {
            if (string.IsNullOrWhiteSpace(FormatCell.Text))
                return string.Empty;

            CSheetFormat Format = FormatCell;

            if (Format.dwFlags.HasFlag(MoxelCellFlags.FontName))
            {
                FontFamily = FormatCell.Parent.FontList[Format.wFontNumber].lfFaceName;
                CellStyle.Set("font-family", CssFontName(FontFamily));
            }

            if (Format.dwFlags.HasFlag(MoxelCellFlags.FontSize))
            {
                FontSize = (float)-Format.wFontSize / 4;
                // Math.Round до целого сохранён как в оригинале (теряет полуторные кегли — осознанно)
                CellStyle.Set("font-size", $"{Inv(Math.Round(FontSize))}pt");
            }

            if (Format.dwFlags.HasFlag(MoxelCellFlags.FontWeight))
                CellStyle.Set("font-weight", Format.bFontBold.ToString());

            if (Format.dwFlags.HasFlag(MoxelCellFlags.FontItalic) && Format.bFontItalic)
                CellStyle.Set("font-style", "italic");

            if (Format.dwFlags.HasFlag(MoxelCellFlags.FontColor))
            {
                int colorIndex = Format.bFontColor;
                if (colorIndex >= 0 && colorIndex < a1CPallete.Length) // ФИКС: было || — всегда true, вылет за границы палитры
                {
                    Color fontColor = Color.FromArgb((int)(a1CPallete[colorIndex] + 0xFF000000));
                    CellStyle.Set("color", Rgb(fontColor));
                }
            }

            bool preventWrap = Format.bControlContent != TextControl.Wrap && !(FormatCell is EmbeddedObject);

            // ФИКС: HTML-экранирование; модель больше не мутирует (раньше FormatCell.Text перезаписывался "&nbsp;")
            string text = System.Net.WebUtility.HtmlEncode(FormatCell.Text);
            if (preventWrap)
            {
                text = text.Replace(" ", "&nbsp;");
                CellStyle.Set("white-space", "nowrap");
                CellStyle.Set("max-width", "0px");
            }

            if (Format.dwFlags.HasFlag(MoxelCellFlags.AlignV))
                CellStyle.Set("vertical-align", Format.bVertAlign.ToString());

            if (Format.dwFlags.HasFlag(MoxelCellFlags.AlignH))
            {
                if (Format.bHorAlign.HasFlag(TextHorzAlign.BySelection) && Format.bHorAlign.HasFlag(TextHorzAlign.Center))
                    CellStyle.Set("text-align", "center");
                else
                {
                    CellStyle.Set("text-align", Format.bHorAlign.ToString());
                    if (Format.bHorAlign == TextHorzAlign.Right)
                        CellStyle.Set("padding-right", "3px");
                }
            }

            if (FormatCell.TextOrientation != 0)
                CellStyle.Set("transform", $"rotate(-{FormatCell.TextOrientation}deg)");

            return text.Replace("\r\n", "<br>");
        }

        public static void FillLineStyle(CSheetFormat FormatCell, ref CSSstyle LineStyle)
        {
            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
            {
                switch (FormatCell.bPictureBorderStyle)
                {
                    case ObjectBorderStyle.DashDotDot: LineStyle.Set("stroke-dasharray", "11 3 3 3 3 3"); break;
                    case ObjectBorderStyle.DashDotSparse: LineStyle.Set("stroke-dasharray", "8 5 3 5"); break;
                    case ObjectBorderStyle.DashedExtraLong: LineStyle.Set("stroke-dasharray", "16 6"); break;
                    case ObjectBorderStyle.DashedShort: LineStyle.Set("stroke-dasharray", "3 3"); break;
                }
            }

            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
                LineStyle.Set("stroke-width", $"{(byte)FormatCell.bPictureBorderWidth * 2 + 1}px");

            LineStyle.Set("stroke", Rgb(FormatCell.BorderColor)); // выставляется всегда, как в оригинале
        }

        public static void FillCellStyle(CSheetFormat FormatCell, ref CSSstyle CellStyle)
        {
            // Одинаковые стили границ со всех сторон -> shorthand "border"
            bool sameBorder = FormatCell.bBorderTop == FormatCell.bBorderBottom
                           && FormatCell.bBorderBottom == FormatCell.bBorderLeft
                           && FormatCell.bBorderLeft == FormatCell.bBorderRight
                           && FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop);

            if (sameBorder)
            {
                SetBorder(CellStyle, "border", FormatCell.bBorderTop, FormatCell.bBorderColor);
            }
            else
            {
                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
                    SetBorder(CellStyle, "border-top", FormatCell.bBorderTop, FormatCell.bBorderColor);
                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                    SetBorder(CellStyle, "border-left", FormatCell.bBorderLeft, FormatCell.bBorderColor);
                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                    SetBorder(CellStyle, "border-right", FormatCell.bBorderRight, FormatCell.bBorderColor);
                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderBottom))
                    SetBorder(CellStyle, "border-bottom", FormatCell.bBorderBottom, FormatCell.bBorderColor);
            }

            Color bgColor = FormatCell.BgColor;
            if (bgColor != Color.Empty)
                CellStyle.Set("background-color", Rgb(bgColor));
        }

        private static void SetBorder(CSSstyle style, string property, BorderStyle border, int borderColorIndex)
        {
            string css = GetBorderStyle(border, borderColorIndex);
            if (css != null)
                style.Set(property, css);
        }

        private static string GetBorderStyle(BorderStyle moxelBorder, int borderColorIndex)
        {
            uint borderColor = 0;
            if (borderColorIndex >= 0 && borderColorIndex < a1CPallete.Length) // ФИКС: было || — всегда true
                borderColor = a1CPallete[borderColorIndex];

            string color = $"#{borderColor:X6}";
            switch (moxelBorder)
            {
                case BorderStyle.None: return "#ffffff 0px none";
                case BorderStyle.ThinDotted:
                case BorderStyle.ThinGrayDotted: return $"{color} 1px dotted";
                case BorderStyle.ThinSolid: return $"{color} 1px solid";
                case BorderStyle.MediumSolid: return $"{color} 2px solid";
                case BorderStyle.ThickSolid: return $"{color} 3px solid";
                case BorderStyle.Double: return $"{color} 1px double";
                case BorderStyle.ThinDashedShort:
                case BorderStyle.ThinDashedLong:
                case BorderStyle.MediumDashed: return $"{color} 1px dashed";
                default:
                    return null; // ФИКС: раньше пустая строка давала невалидное "border-top: ;"
            }
        }

        #endregion

        #region SVG-паттерны

        private static string GetSvgFillPattern(CSheetFormat formatCell, string patternId)
        {
            if (!formatCell.dwFlags.HasFlag(MoxelCellFlags.PatternType))
                return string.Empty;

            Color patternColor = formatCell.PatternColor;
            Color bgColor = formatCell.BgColor;
            if (bgColor == Color.Empty)
                bgColor = Color.White;

            string fg = Rgb(patternColor);
            string bg = Rgb(bgColor);

            var result = new StringBuilder();
            // ФИКС: id теперь уникальный (раньше все ячейки ссылались на один "defpattern")
            result.AppendLine($"<style type=\"text/css\">\r\n\t#{patternId} {{ fill: {fg}; }}\r\n\t</style>");

            switch (formatCell.bPatternType)
            {
                case 1:
                    result.AppendLine($"<defs>\r\n\t<pattern id=\"{patternId}\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"{bg}\"/>\r\n\t\t<rect x=\"0\" y=\"1\" width=\"2\" height=\"1\"/>\r\n\t</pattern>\r\n\t</defs>");
                    break;
                case 2:
                    result.AppendLine($"<defs>\r\n\t<pattern id=\"{patternId}\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"{bg}\"/>\r\n\t\t<rect x=\"1\" y=\"1\" width=\"1\" height=\"1\"/>\r\n\t\t<rect x=\"0\" y=\"1\" width=\"1\" height=\"1\" fill=\"{bg}\"/>\r\n\t</pattern>\r\n\t</defs>");
                    break;
                case 3:
                    result.AppendLine($"<defs>\r\n\t<pattern id=\"{patternId}\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"{bg}\"/>\r\n\t\t<rect x=\"0\" y=\"1\" width=\"2\" height=\"1\" fill=\"{bg}\"/>\r\n\t</pattern>\r\n\t</defs>");
                    break;
                case 4:
                    result.AppendLine($"<defs>\r\n\t<pattern id=\"{patternId}\" patternUnits=\"userSpaceOnUse\" width=\"8\" height=\"8\"><rect x=\"0\" y=\"0\" width=\"8\" height=\"8\"/>\r\n\t\t<rect x=\"0\" y=\"0\" width=\"3\" height=\"2\" fill=\"{bg}\"/>\r\n\t\t<rect x=\"4\" y=\"0\" width=\"4\" height=\"2\" fill=\"{bg}\"/>\r\n\t\t<rect x=\"0\" y=\"3\" width=\"7\" height=\"3\" fill=\"{bg}\"/>\r\n\t\t<rect x=\"0\" y=\"7\" width=\"3\" height=\"1\" fill=\"{bg}\"/>\r\n\t\t<rect x=\"4\" y=\"7\" width=\"4\" height=\"1\" fill=\"{bg}\"/>\r\n\t</pattern>\r\n\t</defs>");
                    break;
                case 5:
                    // ФИКС: в stroke был PatternColor.G вместо .B (синий канал заменялся зелёным)
                    result.AppendLine($"<defs>\r\n\t<pattern id=\"{patternId}\" patternUnits=\"userSpaceOnUse\" width=\"8\" height=\"8\"><rect stroke-width=\"0\" height=\"8\" width=\"8\" y=\"0\" x=\"0\" fill=\"{bg}\"></rect><path d=\" M 0 0.5 H 1 M 4 0.5 H 5 M 1 1.5 H 2 M 3 1.5 H 4 M 5 1.5 H 6 M 2 2.5 H 3 M 6 2.5 H 7 M 1 3.5 H 2 M 5 3.5 H 6 M 7 3.5 H 8 M 0 4.5 H 1 M 4 4.5 H 5 M 3 5.5 H 4 M 5 5.5 H 6 M 7 5.5 H 8 M 2 6.5 H 3 M 6 6.5 H 7 M 1 7.5 H 2 M 3 7.5 H 4 M 7 7.5 H 8\" stroke-width=\"1\" stroke=\"{fg}\"/>\r\n\t</pattern>\r\n\t</defs>");
                    break;
            }

            return result.ToString();
        }

        #endregion

        #region Границы картинок

        private static string PictureBorderWidth(ObjectBorderWidth bPictureBorderWidth)
        {
            switch (bPictureBorderWidth)
            {
                case ObjectBorderWidth.Medium: return "2px ";
                case ObjectBorderWidth.Thick: return "3px ";
                case ObjectBorderWidth.Thin:
                default: return "0.5px ";
            }
        }

        private static string PictureBorderStyle(CSheetFormat FormatCell)
        {
            if (FormatCell.bPictureBorderPresence == ObjectBorderPresence.Empty)
                return "none";

            string width = PictureBorderWidth(FormatCell.bPictureBorderWidth);
            switch (FormatCell.bPictureBorderStyle)
            {
                case ObjectBorderStyle.None:
                case ObjectBorderStyle.Solid:
                    return $"{width}solid";
                case ObjectBorderStyle.DashedShort:
                    return $"{width}dotted";
                case ObjectBorderStyle.DashedExtraLong:
                case ObjectBorderStyle.DashDotSparse:
                case ObjectBorderStyle.DashDotDot:
                    return $"{width}dashed";
                default:
                    return "none";
            }
        }

        #endregion



        #region Публичный API

        public static bool Save(Moxel moxel, string filename)
        {
            using (var fs = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None, 64 * 1024))
            {
                RenderToHtml(moxel, fs);
            }
            return File.Exists(filename);
        }

        /// <summary>
        /// Рендер одного встроенного объекта (картинка/текст/линия/прямоугольник).
        /// </summary>
        public static void RenderImage(TextWriter result, EmbeddedObject obj, Dictionary<int, long> rowHeights, bool inUnion)
        {
            var pictureStyle = new CSSstyle();
            Rectangle area = obj.AbsoluteImageArea;
            CSheetFormat formatCell = obj;
            int borderWidth = (int)formatCell.bPictureBorderWidth;

            string patternId = $"moxelPattern{Interlocked.Increment(ref _patternCounter)}";

            pictureStyle.Set("margin-top", $"{Inv(obj.Picture.dwOffsetTop * 0.25)}pt");
            pictureStyle.Set("margin-left", $"{Inv(obj.Picture.dwOffsetLeft * 0.25)}pt");
            pictureStyle.Set("width", $"{area.Width}px");
            pictureStyle.Set("height", $"{area.Height}px");

            bool drawRectangle = obj.Picture.dwType == ObjectType.Rectangle;

            if (obj.Picture.dwType != ObjectType.Line)
            {
                if (formatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight) || formatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                {
                    if ((formatCell.bPictureBorderPresence != ObjectBorderPresence.All) && formatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                    {
                        if (formatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Left))
                            pictureStyle.Set("border-left", PictureBorderStyle(formatCell));
                        if (formatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Right))
                            pictureStyle.Set("border-right", PictureBorderStyle(formatCell));
                        if (formatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Top))
                            pictureStyle.Set("border-top", PictureBorderStyle(formatCell));
                        if (formatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Bottom))
                            pictureStyle.Set("border-bottom", PictureBorderStyle(formatCell));
                    }
                    else if (formatCell.bPictureBorderStyle <= ObjectBorderStyle.Solid && !formatCell.dwFlags.HasFlag(MoxelCellFlags.PatternType))
                    {
                        pictureStyle.Set("border", PictureBorderStyle(formatCell));
                    }
                    else
                    {
                        drawRectangle = true;
                    }
                }
                else
                {
                    pictureStyle.Set("border",
                        formatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop)
                            ? $"solid {(byte)formatCell.bPictureBorderWidth * 2 + 1}px"
                            : "solid 1px");
                }

                if (!drawRectangle)
                    pictureStyle.Set("border-color", Rgb(formatCell.BorderColor));

                Color bgColor = formatCell.BgColor;
                if (bgColor != Color.Empty)
                    pictureStyle.Set("background-color", Rgb(bgColor));
            }

            if (inUnion)
                pictureStyle.Set("position", "absolute");
            pictureStyle.Set("overflow", "hidden");

            if (formatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV) && formatCell.bVertAlign == TextVertAlign.Middle)
            {
                pictureStyle.Set("display", "flex");
                pictureStyle.Set("align-items", "center");
            }

            if (formatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH))
            {
                if (formatCell.bHorAlign.HasFlag(TextHorzAlign.BySelection) && formatCell.bHorAlign.HasFlag(TextHorzAlign.Center))
                    pictureStyle.Set("text-align", "center");
                else
                    pictureStyle.Set("text-align", formatCell.bHorAlign.ToString());
            }

            var lineStyle = new CSSstyle();
            FillLineStyle(formatCell, ref lineStyle); // stroke выставится внутри всегда

            if (drawRectangle)
            {
                lineStyle.Set("fill", formatCell.dwFlags.HasFlag(MoxelCellFlags.PatternType) ? $"url(#{patternId})" : "none");

                var svg = new StringBuilder();
                svg.AppendLine($"<svg baseProfile=\"full\" xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" height = \"{area.Height + borderWidth}px\"  width = \"{area.Width + borderWidth}px\" text-rendering=\"geometricPrecision\">");
                svg.AppendLine(GetSvgFillPattern(formatCell, patternId));
                svg.AppendLine($"<g transform=\"translate({borderWidth / 2}, {borderWidth / 2})\">");
                svg.AppendLine($"<rect {lineStyle} x=\"1\" y=\"1\" width=\"{area.Width}\" height=\"{area.Height}\"/>");
                svg.AppendLine("</g>");
                svg.AppendLine("</svg>");
                pictureStyle.Set("background-image", $"url(data:image/svg+xml;base64,{Convert.ToBase64String(Encoding.UTF8.GetBytes(svg.ToString()))})");
            }

            result.Write($"\t\t<span id=\"D{obj.Picture.dwZOrder}\"{pictureStyle}>\r\n");

            switch (obj.Picture.dwType)
            {
                case ObjectType.Ole:
                case ObjectType.Picture:
                    using (var ms = new MemoryStream())
                    {
                        // Странный косяк GDI+: без пересоздания битмапа Save даёт неопознанную ошибку
                        using (Bitmap bmp = obj.pObject)
                            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

                        result.Write($"\t\t\t<img src=\"data:image/png;base64,{Convert.ToBase64String(ms.ToArray())}\" width=\"{area.Width + borderWidth}\" height=\"{area.Height + borderWidth}\" alt=\"\">\r\n");
                    }
                    break;

                case ObjectType.Text:
                    var textStyle = new CSSstyle();
                    string text = FillTextStyle(obj, ref textStyle);
                    textStyle.Set("max-width", $"{area.Width}px");
                    textStyle.Set("width", $"{area.Width}px");
                    textStyle.Set("line-height", "1.57");
                    result.WriteLine($"\t\t\t<span{textStyle}>{text}</span>");
                    break;

                case ObjectType.Line:
                    result.WriteLine($"<svg baseProfile=\"full\" xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" height = \"{Math.Max(area.Height, 10)}px\"  width = \"{Math.Max(area.Width, 10)}px\" text-rendering=\"geometricPrecision\">");
                    result.WriteLine("<g transform=\"translate(0.5, 0.5)\">");
                    Rectangle lineCoords = obj.ImageArea;
                    if (lineCoords.Height * lineCoords.Width >= 0)
                        result.WriteLine($"<line {lineStyle} x1=\"0\" y1=\"1\" x2=\"{area.Width}\" y2=\"{Math.Max(area.Height, borderWidth)}\"/>");
                    else
                        result.WriteLine($"<line {lineStyle} x1=\"0\" y2=\"1\" x2=\"{area.Width}\" y1=\"{Math.Max(area.Height, borderWidth)}\"/>");
                    result.WriteLine("</g>");
                    result.WriteLine("</svg>");
                    break;
            }

            result.Write("\t\t</span>\r\n");
        }

        public static void RenderToHtml(Moxel moxel, Stream stream)
        {
            var context = new RenderContext(moxel);

            using var result = new StreamWriter(stream, Encoding.UTF8, 64 * 1024, leaveOpen: true);

            string defaultFontFamily = "Arial";
            string bodyFontAttr = string.Empty;
            float defFontSize = DefaultFontSizePt;

            if (moxel.FontList.Count == 1)
            {
                defaultFontFamily = moxel.FontList.Values.First().lfFaceName;
                bodyFontAttr = $" style=\"font-family:{CssFontName(defaultFontFamily)}\"";
            }

            CSheetFormat defFormat = moxel.DefFormat;
            if (defFormat.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                defFontSize = -(float)defFormat.wFontSize / 4;

            WriteDocumentHead(result, defaultFontFamily, defFontSize, bodyFontAttr);

            // Порядок страниц: полоса 1 (все строки) -> полоса 2 (все строки)...
            // это соответствует Excel-опции печати «Down, then over».
            int totalWork = Math.Max(1, context.ColumnPages.Count * moxel.nAllRowCount);

            for (int p = 0; p < context.ColumnPages.Count; p++)
            {
                var page = context.ColumnPages[p];
                RenderTable(moxel, context, result, page.Start, page.End,
                    pageBreakBefore: p > 0,
                    defaultFontFamily, defFontSize,
                    progressOffset: p * moxel.nAllRowCount,
                    progressTotal: totalWork);
            }

            result.Write("\t</body>\r\n");
            result.Write("</html>");
        }

        private static void RenderTable(Moxel moxel, RenderContext context, TextWriter result,
            int colStart, int colEnd, bool pageBreakBefore,
            string defaultFontFamily, float defFontSize,
            int progressOffset, int progressTotal)
        {
            string tableBreak = pageBreakBefore ? "page-break-before: always; " : string.Empty;
            double tableWidth = Math.Round(moxel.GetWidth(colStart, colEnd + 1) * ColumnWidthToPixels);
            result.Write($"\t\t<TABLE style=\"{tableBreak}width: {Inv(tableWidth)}px; height: 0px;\" border=\"0\" CELLSPACING=\"0\">\r\n");
            result.Write("\t\t\t<colgroup>\r\n");

            for (int column = colStart; column <= colEnd; column++)
            {
                CSheetFormat colFormat = moxel.Columns.TryGetValue(column, out var cf) ? cf : moxel.DefFormat;
                double width = colFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth)
                    ? Math.Round(colFormat.wWidth * ColumnWidthToPixels)
                    : DefaultColumnWidthPx;
                result.Write($"\t\t\t\t<col width=\"{Inv(width)}\"/>\r\n");
            }

            result.Write("\t\t\t</colgroup>\r\n\t\t\t<tbody>\r\n");

            for (int rownumber = 0; rownumber < moxel.nAllRowCount; rownumber++)
            {
                onProgress?.Invoke((progressOffset + rownumber + 1) * 100 / progressTotal);

                moxel.Rows.TryGetValue(rownumber, out MoxelRow row);
                CSheetFormat rowFormat = row?.FormatCell ?? moxel.DefFormat;

                bool rowAutoHeight = true;
                int rowHeight = 0;
                if (row != null)
                {
                    rowAutoHeight =
                        row.All(x => !(x.Value.FormatCell.dwFlags.HasFlag(MoxelCellFlags.RowHeight) && x.Value.FormatCell.wHeight > 0)) &&
                        !(rowFormat.dwFlags.HasFlag(MoxelCellFlags.RowHeight) && rowFormat.wHeight > 0);
                    rowHeight = row.Height;
                }

                List<CellsUnion> rowUnions = context.UnionsByRow.TryGetValue(rownumber, out List<CellsUnion> ul) ? ul : NoUnions;

                // Межстрочные объединения, попадающие в текущую полосу колонок
                HashSet<int> spannedColumns = null;
                foreach (CellsUnion u in rowUnions)
                {
                    if (u.dwTop == u.dwBottom)
                        continue;
                    for (int cc = Math.Max(u.dwLeft, colStart); cc <= Math.Min(u.dwRight, colEnd); cc++)
                        (spannedColumns ??= new HashSet<int>()).Add(cc);
                }

                var rowStyle = new CSSstyle();
                var rowString = new StringBuilder();
                var rowClass = string.Empty;
                // === Горизонтальный разрыв страницы: разрыв ПОСЛЕ этой строки ===
                // Дублируем -webkit-* для старых сборок wkhtmltopdf на Qt WebKit.
                // Если разрыв в вашем формате означает «строка b — первая на новой странице»,
                // ставьте вместо этого page-break-before на троке b.
                if (context.RowBreaks.Contains(rownumber))
                {
                    //rowStyle.Set("page-break-before", "always");
                    //rowStyle.Set("-webkit-page-break-before", "always");

                    result.WriteLine("\t\t\t<tr class=\"forced-break\"></tr>");
                }
             

                for (int columnnumber = colStart; columnnumber <= colEnd; columnnumber++)
                {
                    int c = columnnumber;
                    string text = string.Empty;
                    var cellStyle = new CSSstyle();
                    string fontFamily = defaultFontFamily;
                    float fontSize = defFontSize;
                    CSheetFormat formatCell = moxel.DefFormat;

                    // Объединения обрезаем по границе полосы (struct — работаем с копией)
                    CellsUnion union = default;
                    foreach (CellsUnion u in rowUnions)
                    {
                        int left = Math.Max(u.dwLeft, colStart);
                        int right = Math.Min(u.dwRight, colEnd);
                        if (left <= c && c <= right)
                        {
                            union = u;
                            union.dwLeft = left;
                            union.dwRight = right;
                            break;
                        }
                    }

                    if (row != null)
                    {
                        DataCell cell = row[c];
                        text = FillTextStyle(cell, ref cellStyle, ref fontFamily, ref fontSize);
                        formatCell = cell;
                    }

                    FillCellStyle(formatCell, ref cellStyle);

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        // Сосед справа учитывается только внутри текущей полосы
                        if (c < colEnd)
                        {
                            DataCell nextCell = row[c + 1];

                            if (formatCell.bControlContent == TextControl.Auto)
                            {
                                if (!string.IsNullOrEmpty(nextCell.Text) ||
                                    (nextCell.FormatCell.bBorderLeft != BorderStyle.None && nextCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft)))
                                {
                                    cellStyle.Set("overflow", "hidden");
                                }
                            }

                            // Псевдо-объединения ограничиваем правым краем полосы
                            if (union.IsEmpty() && string.IsNullOrEmpty(nextCell.Text))
                            {
                                if (formatCell.bHorAlign.HasFlag(TextHorzAlign.BySelection) &&
                                    formatCell.bHorAlign.HasFlag(TextHorzAlign.Center) &&
                                    !formatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight) &&
                                    !nextCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                                {
                                    union = new CellsUnion { dwTop = rownumber, dwBottom = rownumber, dwLeft = c, dwRight = colEnd };
                                }
                                else if (formatCell.bControlContent == TextControl.Wrap && !formatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                                {
                                    int cn = c + 1;
                                    while (cn <= colEnd)
                                    {
                                        DataCell probe = row[cn];
                                        if (!string.IsNullOrEmpty(probe.Text) || probe.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                                            break;
                                        cn++;
                                    }

                                    if (cn > c + 1)
                                        union = new CellsUnion { dwTop = rownumber, dwBottom = rownumber, dwLeft = c, dwRight = cn - 1 };
                                }
                            }
                        }

                        if (rowAutoHeight && spannedColumns?.Contains(c) != true)
                        {
                            var fontStyle = formatCell.bFontBold == clFontWeight.Bold ? FontStyle.Bold : FontStyle.Regular;
                            string fontHash = $"{fontFamily}_{fontSize}_{fontStyle}";
                            if (!context.FontCache.TryGetValue(fontHash, out Font font))
                            {
                                font = new Font(fontFamily, (float)(fontSize * 0.95), fontStyle, GraphicsUnit.Point);
                                context.FontCache[fontHash] = font;
                            }

                            string textForRender = text.Replace("&nbsp;", " ").Replace("<br>", "\r\n");
                            var constr = new Size
                            {
                                Width = (int)Math.Round(moxel.GetWidth(c, c + union.ColumnSpan + 1) * MeasureWidthScale),
                                Height = 0
                            };

                            Size textSize = System.Windows.Forms.TextRenderer.MeasureText(
                                textForRender, font, constr,
                                System.Windows.Forms.TextFormatFlags.SingleLine |
                                System.Windows.Forms.TextFormatFlags.Left |
                                System.Windows.Forms.TextFormatFlags.TextBoxControl);

                            string[] lines = textForRender.Split('\r');
                            float charWidth = (float)textSize.Width / textForRender.Length;
                            int lineCount = lines.Length;

                            if (formatCell.bControlContent == TextControl.Wrap)
                            {
                                cellStyle.Set("line-height", "1.25");
                                foreach (string line in lines)
                                {
                                    if (line.Length * charWidth <= constr.Width)
                                        continue;

                                    float lineLen = 0f;
                                    foreach (string word in line.Split(' '))
                                    {
                                        lineLen += (word.Length + 1) * charWidth;
                                        if (lineLen >= constr.Width)
                                        {
                                            lineLen = (word.Length + 1) * charWidth;
                                            lineCount++;
                                        }
                                    }
                                }
                            }

                            if (formatCell.bControlContent == TextControl.Wrap || lineCount > 1)
                            {
                                constr.Height = (int)Math.Ceiling((float)textSize.Width / constr.Width) * textSize.Height;
                                if (lineCount > constr.Height / textSize.Height)
                                    constr.Height = textSize.Height * lineCount + 1;
                            }
                            else
                            {
                                constr.Height = textSize.Height;
                            }

                            constr.Height = Math.Max(15, constr.Height);
                            rowHeight = Math.Max(rowHeight, Math.Min(short.MaxValue, constr.Height * QuarterPointsPerPixel));
                        }

                        if (c > colStart)
                        {
                            DataCell prevCell = row[c - 1];
                            if (string.IsNullOrEmpty(prevCell.Text) &&
                                formatCell.bHorAlign == TextHorzAlign.Right &&
                                !prevCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight) &&
                                !formatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                            {
                                cellStyle.Set("direction", "rtl");
                                cellStyle.Set("overflow", "visible");
                                text = $"<SPAN style=\"white-space: nowrap; direction: ltr; display: inline-block;\">{text}</SPAN>";
                            }
                        }
                    }
                    else if (context.ObjectsByCell.TryGetValue((rownumber, c), out EmbeddedObject pic))
                    {
                        var sb = new StringBuilder();
                        using (TextWriter tw = new StringWriter(sb))
                            RenderImage(tw, pic, context.RowHeights, spannedColumns?.Contains(c) != true);
                        text = sb.ToString();
                        cellStyle.Set("align-content", "flex-start");
                    }

                    if (!union.ContainsCell(rownumber, c))
                    {
                        if (!string.IsNullOrWhiteSpace(text))
                            rowString.Append($"\t\t\t\t<td{union.HtmlSpan}{cellStyle}>{text}</td>\r\n");
                        else
                            rowString.Append($"\t\t\t\t<td{union.HtmlSpan}{cellStyle}></td>\r\n");
                    }

                    columnnumber += union.ColumnSpan;
                }

                // Высота строки: явная > авто > дефолт. Для согласованности сетки строк
                // между полосами берём максимум с рассчитанным в предыдущих полосах.
                double? heightPt = null;
                if (row != null && rowFormat.dwFlags.HasFlag(MoxelCellFlags.RowHeight) && rowFormat.wHeight > 0)
                {
                    heightPt = Math.Round(rowFormat.wHeight * QuarterPointToPoints, 2);
                    context.RowHeights[rownumber] = rowFormat.wHeight;
                }
                else
                {
                    context.RowHeights.TryGetValue(rownumber, out long stored);
                    rowHeight = Math.Max(rowHeight, (int)stored);

                    if (rowHeight > 0)
                    {
                        heightPt = Math.Round(rowHeight * QuarterPointToPoints, 2);
                        context.RowHeights[rownumber] = rowHeight;
                    }
                    else
                    {
                        context.RowHeights[rownumber] = DefaultRowHeightQP;
                    }
                }

                if (heightPt.HasValue)
                    rowStyle.Set("height", $"{Inv(heightPt.Value)}pt");

                result.Write($"\t\t\t<tr{rowStyle} {rowClass}id=\"R{rownumber:00}\">\r\n{rowString}\t\t\t</tr>\r\n");
            }

            result.Write("\t\t</table>\r\n");
        }

        private static void WriteDocumentHead(TextWriter result, string fontFamily, float fontSize, string bodyFontAttr)
        {
            // ВНИМАНИЕ: DOCTYPE оставлен как в оригинале (quirks mode, под него откалиброваны коэффициенты).
            result.Write("<!DOCTYPE HTML PUBLIC \" -//W3C//DTD HTML 5.0 Transitional//EN\">\r\n<HTML>\r\n");
            result.Write("<HEAD>\r\n<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html; CHARSET=utf-8\"/>\r\n");
            result.Write("<style type=\"text/css\">\r\n");
            result.Write("body { background: #ffffff; margin: 0; font-family: Arial; font-size: 8pt; font-style: normal; }\r\n");
            result.Write($"table {{ table-layout: fixed; padding: 0px; padding-left: 2px; vertical-align: bottom; border-collapse: collapse; width: 100%; font-family: {CssFontName(fontFamily)}; font-size: {Inv(fontSize)}pt; font-style: normal; }}\r\n");
            result.Write("td { padding: 0px 0px 0px 1px; }\r\n");
            result.Write(".forced-break { display: block; page-break-before: always; height: 1px; }\r\n");
            // page-break-inside: avoid — не резать строку между страницами при экспорте в PDF
            result.Write("tr { height: 11.25pt; page-break-inside: avoid; }\r\n");
            result.Write("</style>\r\n");
            result.Write("</HEAD>\r\n");
            result.Write($"\t<body{bodyFontAttr}>\r\n");
        }
        #endregion

        #region Предрасчитанный контекст рендеринга

        private sealed class RenderContext
        {
            public readonly Dictionary<int, List<CellsUnion>> UnionsByRow;
            public readonly Dictionary<(int Row, int Column), EmbeddedObject> ObjectsByCell;
            public readonly Dictionary<string, Font> FontCache = new Dictionary<string, Font>();
            public readonly Dictionary<int, long> RowHeights = new Dictionary<int, long>();

            /// <summary>Строки, ПОСЛЕ которых идёт разрыв страницы.</summary>
            public readonly HashSet<int> RowBreaks;

            /// <summary>Вертикальные полосы (диапазоны колонок). Одна таблица = одна полоса.</summary>
            public readonly List<(int Start, int End)> ColumnPages;

            public RenderContext(Moxel moxel)
            {
                UnionsByRow = new Dictionary<int, List<CellsUnion>>();
                foreach (CellsUnion u in moxel.Unions)
                {
                    for (int r = u.dwTop; r <= u.dwBottom; r++)
                    {
                        if (!UnionsByRow.TryGetValue(r, out List<CellsUnion> list))
                            UnionsByRow[r] = list = new List<CellsUnion>();
                        list.Add(u);
                    }
                }

                ObjectsByCell = new Dictionary<(int Row, int Column), EmbeddedObject>();
                foreach (EmbeddedObject obj in moxel.Objects)
                {
                    var key = (obj.Picture.dwRowStart, obj.Picture.dwColumnStart);
                    if (!ObjectsByCell.ContainsKey(key))
                        ObjectsByCell[key] = obj;
                }

                // Разрыв внутри вертикального объединения порезал бы rowspan — отфильтровываем.
                RowBreaks = new HashSet<int>(
                    (moxel.HorisontalPageBreaks ?? Array.Empty<int>())
                        .Where(b => b >= 0 && b < moxel.nAllRowCount - 1)
                        .Where(b => !moxel.Unions.Any(u => u.dwTop <= b && b < u.dwBottom)));

                ColumnPages = GetColumnPages(moxel.VerticalPageBreaks, moxel.nAllColumnCount);
            }

            /// <summary>
            /// Разбивает колонки на полосы по VerticalPageBreaks.
            /// Семантика: разрыв ПОСЛЕ колонки b (b — последняя колонка страницы).
            /// Если в вашем формате разрыв означает «ПЕРЕД колонкой b» — замените тело цикла на:
            ///     int end = b - 1;
            ///     if (end >= start) { pages.Add((start, end)); start = b; }
            /// </summary>
            private static List<(int Start, int End)> GetColumnPages(int[] breaks, int columnCount)
            {
                var pages = new List<(int Start, int End)>();
                int start = 0;

                if (breaks != null)
                    foreach (int b in breaks.Distinct().OrderBy(x => x))
                    {
                        int end = Math.Min(b, columnCount - 1);
                        if (end >= start)
                        {
                            pages.Add((start, end));
                            start = end + 1;
                        }
                    }

                if (start <= columnCount - 1)
                    pages.Add((start, columnCount - 1));

                if (pages.Count == 0)
                    pages.Add((0, Math.Max(0, columnCount - 1)));

                return pages;
            }
        }

        #endregion
    }
}