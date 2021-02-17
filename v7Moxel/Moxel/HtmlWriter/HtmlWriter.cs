using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using static Moxel.Moxel;
using System.IO;
//using DocumentFormat.OpenXml.Office2010.Excel;

namespace Moxel
{
    public class HtmlWriter
    {

        public static event ConverterProgressor onProgress;

        public class CSSstyle : Dictionary<string, string>
        {

            public override string ToString()
            {
                if (Count > 0)
                    return $" style=\"{ string.Join("; ", this.Select(t => t.Key + ": " + t.Value))}\"";
                else
                    return string.Empty;
            }

            internal void Set(string key, string value)
            {
                if (ContainsKey(key))
                    this[key] = value;
                else
                    Add(key, value);
            }
        }

        public static string FillTextStyle(DataCell FormatCell, ref CSSstyle CellStyle)
        {
            string FontFamily = string.Empty;
            float FontSize = 0;
            return FillTextStyle(FormatCell, ref CellStyle, ref FontFamily, ref FontSize);
        }

        public static string FillTextStyle(DataCell FormatCell, ref CSSstyle CellStyle, ref string FontFamily, ref float FontSize)
        {
            CSheetFormat Format = FormatCell;
            string Text = string.Empty;

            if (!string.IsNullOrWhiteSpace(FormatCell.Text))
            {
                if (Format.dwFlags.HasFlag(MoxelCellFlags.FontName))
                {
                    FontFamily = FormatCell.Parent.FontList[Format.wFontNumber].lfFaceName;
                    CellStyle.Set("font-family", FontFamily);
                }

                if (Format.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                {
                    FontSize = (float)-Format.wFontSize / 4;
                    CellStyle.Set("font-size", $"{Math.Round(FontSize, 0)}pt");
                }

                if (Format.dwFlags.HasFlag(MoxelCellFlags.FontWeight))
                    CellStyle.Set("font-weight", Format.bFontBold.ToString());

                if (Format.dwFlags.HasFlag(MoxelCellFlags.FontItalic))
                    if (Format.bFontItalic)
                        CellStyle.Set("font-style", "italic");

                if (Format.dwFlags.HasFlag(MoxelCellFlags.FontColor))
                {
                    int ColorIndex = Format.bFontColor;
                    if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                    {
                        Color bgColor = Color.FromArgb((int)(a1CPallete[ColorIndex] + 0xFF000000));
                        CellStyle.Set("color", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");
                    }
                }

                if (Format.bControlContent != TextControl.Wrap && !(FormatCell is EmbeddedObject))
                {
                    FormatCell.Text = FormatCell.Text.Replace(" ", "&nbsp;");
                    CellStyle.Set("white-space", "nowrap");
                    CellStyle.Set("max-width", "0px");

                }

                if (Format.dwFlags.HasFlag(MoxelCellFlags.AlignV))
                    CellStyle.Set("vertical-align", Format.bVertAlign.ToString());

                if (Format.dwFlags.HasFlag(MoxelCellFlags.AlignH))
                {
                    if (Format.bHorAlign.HasFlag( TextHorzAlign.BySelection) && Format.bHorAlign.HasFlag(TextHorzAlign.Center))
                    {
                        CellStyle.Set("text-align", "center");
                    }
                    else
                        CellStyle.Set("text-align", Format.bHorAlign.ToString());
                }
                if (FormatCell.TextOrientation != 0)
                    CellStyle.Set("transform", $"rotate(-{FormatCell.TextOrientation}deg)");

                Text = FormatCell.Text.Replace("\r\n", "<br>");
            }

            return Text;
        }

        public static void FillLineStyle(CSheetFormat FormatCell, ref CSSstyle LineStyle)
        {
            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                switch (FormatCell.bPictureBorderStyle)
                {
                    case ObjectBorderStyle.DashDotDot:
                        LineStyle.Set("stroke-dasharray", "11 3 3 3 3 3");
                        break;
                    case ObjectBorderStyle.DashDotSparse:
                        LineStyle.Set("stroke-dasharray", "8 5 3 5");
                        break;
                    case ObjectBorderStyle.DashedExtraLong:
                        LineStyle.Set("stroke-dasharray", "16 6");
                        break;
                    case ObjectBorderStyle.DashedShort:
                        LineStyle.Set("stroke-dasharray", "3 3");
                        break;
                    case ObjectBorderStyle.Solid:
                        break;
                    default:
                        break;
                }

            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
            {
                LineStyle.Set("stroke-width", $"{(byte)FormatCell.bPictureBorderWidth * 2 + 1}px");
            }

            Color bgColor = FormatCell.BorderColor;
            LineStyle.Set("stroke", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");
        }

        public static void FillCellStyle(CSheetFormat FormatCell, ref CSSstyle CellStyle)
        {
            if (FormatCell.bBorderTop == FormatCell.bBorderBottom
                && FormatCell.bBorderBottom == FormatCell.bBorderLeft
                && FormatCell.bBorderLeft == FormatCell.bBorderRight
                && FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
                CellStyle.Set("border", GetBorderStyle(FormatCell.bBorderTop, FormatCell.bBorderColor));
            else
            {
                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
                    CellStyle.Set("border-top", GetBorderStyle(FormatCell.bBorderTop, FormatCell.bBorderColor));

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                    CellStyle.Set("border-left", GetBorderStyle(FormatCell.bBorderLeft, FormatCell.bBorderColor));

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                    CellStyle.Set("border-right", GetBorderStyle(FormatCell.bBorderRight, FormatCell.bBorderColor));

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderBottom))
                    CellStyle.Set("border-bottom", GetBorderStyle(FormatCell.bBorderBottom, FormatCell.bBorderColor));
            }

            Color bgColor = FormatCell.BgColor;
            if (bgColor != Color.Empty)
                CellStyle.Set("background-color", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");

        }

        static string GetBorderStyle(BorderStyle Moxelborder, int borderColorIndex)
        {
            uint borderColor = 0;

            if (borderColorIndex >= 0 || borderColorIndex < a1CPallete.Length)
                borderColor = a1CPallete[borderColorIndex];

            switch (Moxelborder)
            {
                case BorderStyle.None:
                    return $"#ffffff 0px none";
                case BorderStyle.ThinDotted:
                    return $"#{borderColor:X6} 1px dotted";
                case BorderStyle.ThinGrayDotted:
                    return $"#{borderColor:X6} 1px dotted";
                case BorderStyle.ThinSolid:
                    return $"#{borderColor:X6} 1px solid";
                case BorderStyle.MediumSolid:
                    return $"#{borderColor:X6} 2px solid";
                case BorderStyle.ThickSolid:
                    return $"#{borderColor:X6} 3px solid";
                case BorderStyle.Double:
                    return $"#{borderColor:X6} 1px double";
                case BorderStyle.ThinDashedShort:
                    return $"#{borderColor:X6} 1px dashed";
                case BorderStyle.ThinDashedLong:
                    return $"#{borderColor:X6} 1px dashed";
                case BorderStyle.MediumDashed:
                    return $"#{borderColor:X6} 1px dashed";
                default:
                    return string.Empty;
            }
        }

        static string GetSVGFilPattern(CSheetFormat formatCell)
        {
            Color PatternColor = formatCell.PatternColor;
            StringBuilder result = new StringBuilder($"<style type=\"text/css\">\r\n\t#defpattern {{ fill: rgb({PatternColor.R},{PatternColor.G}, {PatternColor.B}); }}\r\n\t</style>");
            Color bgColor = formatCell.BgColor;
            if (bgColor == Color.Empty)
                bgColor = Color.White;

            string bgFill = $"rgb({bgColor.R},{bgColor.G}, {bgColor.B})";



            if (formatCell.dwFlags.HasFlag(MoxelCellFlags.PatternType))
            {
                switch (formatCell.bPatternType)
                {
                    case 1:
                        result.AppendLine($"<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"{bgFill}\"/>\r\n\t\t<rect x=\"0\" y=\"1\" width=\"2\" height=\"1\"></rect>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    case 2:
                        result.AppendLine($"<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"{bgFill}\"/>\r\n\t\t<rect x=\"1\" y=\"1\" width=\"1\" height=\"1\"></rect>\r\n\t\t\t<rect x=\"0\" y=\"1\" width=\"1\" height=\"1\" fill=\"{bgFill}\"/>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    case 3:
                        result.AppendLine($"<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"{bgFill}\"/>\r\n\t\t<rect x=\"0\" y=\"1\" width=\"2\" height=\"1\" fill=\"{bgFill}\"/>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    case 4:
                        result.AppendLine($"<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"8\" height=\"8\"><rect x=\"0\" y=\"0\" width=\"8\" height=\"8\"/>\r\n\t\t<rect x=\"0\" y=\"0\" width=\"3\" height=\"2\" fill=\"{bgFill}\"/>\r\n\t\t<rect x=\"4\" y=\"0\" width=\"4\" height=\"2\" fill=\"{bgFill}\"/>\r\n\t\t<rect x=\"0\" y=\"3\" width=\"7\" height=\"3\" fill=\"{bgFill}\"/>\r\n\t\t<rect x=\"0\" y=\"7\" width=\"3\" height=\"1\" fill=\"{bgFill}\"/>\r\n\t\t<rect x=\"4\" y=\"7\" width=\"4\" height=\"1\" fill=\"{bgFill}\"/>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    case 5:
                        result.AppendLine($"<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"8\" height=\"8\"><rect stroke-width=\"0\" height=\"8\" width=\"8\" y=\"0\" x=\"0\" fill=\"{bgFill}\"></rect><path d=\" M 0 0.5 H 1 M 4 0.5 H 5 M 1 1.5 H 2 M 3 1.5 H 4 M 5 1.5 H 6 M 2 2.5 H 3 M 6 2.5 H 7 M 1 3.5 H 2 M 5 3.5 H 6 M 7 3.5 H 8 M 0 4.5 H 1 M 4 4.5 H 5 M 3 5.5 H 4 M 5 5.5 H 6 M 7 5.5 H 8 M 2 6.5 H 3 M 6 6.5 H 7 M 1 7.5 H 2 M 3 7.5 H 4 M 7 7.5 H 8\" stroke-width=\"1\" stroke=\"rgb({PatternColor.R},{PatternColor.G}, {PatternColor.G})\"/>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    default:
                        break;
                }
            }
            else
                return string.Empty;

            return result.ToString();
        }

        public static bool Save(Moxel moxel, string filename)
        {
            using (var fs =  new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None, 1024 * 1024 * 10))
            {
                RenderToHtml(moxel, fs);
                fs.Flush();
            }
            return File.Exists(filename);
        }

        public static void RenderToHtml(Moxel moxel, Stream stream)
        {
            using (var result = new StreamWriter(stream, Encoding.UTF8, 1024 * 1024 * 10, true))
            {

                CSheetFormat FormatCell = moxel.DefFormat;
                string DefFontName = string.Empty;
                float DefFontSize = 8.0f;


                if (moxel.FontList.Count == 1)
                    DefFontName = $" style=\"font-family:{moxel.FontList.First().Value.lfFaceName}\"";

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                    DefFontSize = -(float)FormatCell.wFontSize / 4;

                result.WriteLine("<!DOCTYPE HTML PUBLIC \" -//W3C//DTD HTML 5.0 Transitional//EN\">\r\n<HTML>\r\n");
                result.WriteLine("<HEAD>\r\n<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html; CHARSET = utf-8\"/>");

                result.WriteLine("<style type=\"text/css\">");
                result.WriteLine("body { background: #ffffff; margin: 0; font-family: Arial; font-size: 8pt; font-style: normal; }");
                result.WriteLine("table {table-layout: fixed; padding: 0px; padding-left: 2px; vertical-align:bottom; border-collapse:collapse;width: 100%; font-family: Arial; font-size: 8pt; font-style: normal; }");
                result.WriteLine("td { padding: 0px 0px 0px 2px;}");
                result.WriteLine("tr { height: 15px;}");
                result.WriteLine("</style>");

                result.WriteLine("</HEAD>");
                result.WriteLine($"\t<body{DefFontName}>");
                result.WriteLine($"\t\t<TABLE style=\"width: {Math.Round(moxel.GetWidth(0, moxel.nAllColumnCount) * 0.875) }px; height: 0px; \" border=0 CELLSPACING=0>");
                result.WriteLine($"\t\t\t<colgroup>");

                for (int columnnumber = 0; columnnumber < moxel.nAllColumnCount; columnnumber++)
                {

                    string Width;

                    if (moxel.Columns.ContainsKey(columnnumber))
                        FormatCell = moxel.Columns[columnnumber];
                    else
                        FormatCell = moxel.DefFormat;

                    if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                        Width = $" width={Math.Round(FormatCell.wWidth * 0.875).ToString()}";
                    else
                        Width = " width=35";

                    result.WriteLine($"\t\t\t\t<col{Width}/>");
                }

                result.Write($"\t\t\t</colgroup>\r\n");
                result.Write($"\t\t\t<tbody>\r\n");

                for (int rownumber = 0; rownumber < moxel.nAllRowCount; rownumber++)
                {
                    int progress = (rownumber + 1) * 100 / moxel.nAllRowCount;
                    onProgress?.Invoke(progress);

                    MoxelRow Row = null;
                    FormatCell = moxel.DefFormat;
                    CSSstyle RowStyle = new CSSstyle();
                    StringBuilder RowString = new StringBuilder();

                    if (moxel.Rows.ContainsKey(rownumber))
                    {
                        Row = moxel.Rows[rownumber];
                        FormatCell = Row.FormatCell;
                    }

                    List<CellsUnion> RowUnion = moxel.Unions.Where(t => (t.dwTop <= rownumber && t.dwBottom >= rownumber)).ToList();

                    for (int columnnumber = 0; columnnumber < moxel.nAllColumnCount; columnnumber++)
                    {
                        string Text = string.Empty;
                        CSSstyle CellStyle = new CSSstyle();
                        CellsUnion Union = RowUnion.FirstOrDefault(t => (t.dwLeft <= columnnumber && t.dwRight >= columnnumber));
                        var c = columnnumber;
                        string FontFamily = DefFontName;
                        float FontSize = DefFontSize;

                        if (Row != null)
                        {
                            Text = FillTextStyle(Row[c], ref CellStyle, ref FontFamily, ref FontSize);
                            FormatCell = Row[c];
                        }

                        FillCellStyle(FormatCell, ref CellStyle);
                        var NextColumnCelll = Row?[c + 1];

                        if (!string.IsNullOrWhiteSpace(Text))
                        {
                            if (FormatCell.bControlContent == TextControl.Auto)
                            {
                                if (columnnumber < Row.Count - 1)
                                {

                                    if (!string.IsNullOrEmpty(NextColumnCelll.Text) || NextColumnCelll.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                                    {
                                        CellStyle.Add("Overflow", "Hidden");
                                    }
                                }
                            }

                            if (Union.IsEmpty())
                            {

                                if (string.IsNullOrEmpty(NextColumnCelll.Text))
                                {
                                    if (FormatCell.bHorAlign.HasFlag(TextHorzAlign.BySelection) && FormatCell.bHorAlign.HasFlag(TextHorzAlign.Center)
                                        && !FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight)
                                        && !NextColumnCelll.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                                    {
                                        Union = new CellsUnion { dwTop = rownumber, dwBottom = rownumber, dwLeft = 0, dwRight = moxel.nAllColumnCount };
                                    }
                                    else if (FormatCell.bControlContent == TextControl.Wrap && !FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                                    {
                                        var cn = c + 1;

                                        while (string.IsNullOrEmpty(Row[cn].Text) && cn < moxel.nAllColumnCount && !Row[cn].FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                                            cn++;

                                        if (cn > columnnumber + 1)
                                            Union = new CellsUnion { dwTop = rownumber, dwBottom = rownumber, dwLeft = c, dwRight = cn };
                                    }
                                }

                            }
                            //if (Row.Height == 0)
                            {
                                Size Constr = new Size { Width = (int)Math.Round((moxel.GetWidth(c, c + Union.ColumnSpan) + moxel.GetColumnWidth(c)) * 0.875), Height = 0 };
                                Size textsize = System.Windows.Forms.TextRenderer.MeasureText(Text.TrimStart(' '), new Font(FontFamily, FontSize, FormatCell.bFontBold == clFontWeight.Bold ? FontStyle.Bold : FontStyle.Regular), Constr, System.Windows.Forms.TextFormatFlags.WordBreak);
                                textsize.Height /= Union.RowSpan + 1;

                                if (FormatCell.bControlContent == TextControl.Wrap && !Text.Contains("<br>"))
                                    if (textsize.Width > Constr.Width)
                                    {
                                        int index = (textsize.Width - Constr.Width) / (textsize.Width / Text.Length);
                                        if (index > 1 && index < Text.Length)
                                            Text = Text.Insert(Text.Length - index, "<br>");
                                    }

                                Row.Height = Math.Max((short)(Math.Max(textsize.Height, 15) * 3), Row.Height);
                            }


                            if (c > 0 && string.IsNullOrEmpty(Row[c - 1].Text)
                                && FormatCell.bHorAlign == TextHorzAlign.Right
                                && !Row[c - 1].FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight)
                                && !FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                            {
                                CellStyle.Set("direction", "rtl");
                                CellStyle.Set("Overflow", "Visible");
                                Text = $"<SPAN style=\"white-space: nowrap; direction: ltr; display: inline-block;\">{Text}</SPAN>";
                            }
                        }



                        if (!Union.ContainsCell(rownumber, c))
                            if (!string.IsNullOrWhiteSpace(Text))
                                RowString.AppendLine($"\t\t\t\t<td{Union.HtmlSpan}{CellStyle}>{Text}</td>");
                            else
                                RowString.AppendLine($"\t\t\t\t<td{Union.HtmlSpan}{CellStyle}/>");

                        columnnumber += Union.ColumnSpan;
                    }

                    if (Row != null && Row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                    {
                        RowStyle.Set("height", $"{Row.FormatCell.wHeight / 3}px");
                    }
                    result.WriteLine($"\t\t\t<tr{RowStyle} id=\"R{rownumber:00}\">\r\n{RowString}\r\n\t\t\t</tr>");

                    if (Row != null)
                        Row.Height = Math.Max(45, Row.Height);
                }

                result.WriteLine($"\t\t\t</tbody>");
                result.WriteLine("\t\t</table>");
                foreach (EmbeddedObject obj in moxel.Objects)
                {
                    CSSstyle PictureStyle = new CSSstyle();


                    Rectangle DrawingArea = obj.AbsoluteImageArea;

                    FormatCell = obj;

                    int BorderWith = (int)FormatCell.bPictureBorderWidth * 2 + 1;


                    PictureStyle.Add("top", $"{DrawingArea.Top - BorderWith}px");
                    PictureStyle.Add("left", $"{DrawingArea.Left - BorderWith}px");

                    PictureStyle.Add("width", $"{DrawingArea.Width + BorderWith}px");
                    PictureStyle.Add("height", $"{DrawingArea.Height + BorderWith}px");
                    string Text = string.Empty;

                    bool DrawRectangle = obj.Picture.dwType == ObjectType.Rectangle;

                    if (obj.Picture.dwType != ObjectType.Line)
                    {
                        if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight) || FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                        {
                            if ((FormatCell.bPictureBorderPresence != ObjectBorderPresence.All) && FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                            {
                                if (FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Left))
                                    PictureStyle.Set("border-left", PictureBorderStyle(FormatCell));

                                if (FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Right))
                                    PictureStyle.Set("border-right", PictureBorderStyle(FormatCell));

                                if (FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Top))
                                    PictureStyle.Set("border-top", PictureBorderStyle(FormatCell));

                                if (FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Bottom))
                                    PictureStyle.Set("border-bottom", PictureBorderStyle(FormatCell));
                            }
                            else
                                if (FormatCell.bPictureBorderStyle <= ObjectBorderStyle.Solid && !FormatCell.dwFlags.HasFlag(MoxelCellFlags.PatternType))
                            {
                                PictureStyle.Set("border", PictureBorderStyle(FormatCell));
                            }
                            else
                                DrawRectangle = true;
                        }
                        else
                        {
                            if (!FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
                                PictureStyle.Set("border", "solid 1px");
                            else
                                PictureStyle.Set("border", $"solid {(byte)FormatCell.bPictureBorderWidth * 2 + 1}px");
                        }

                        if (!DrawRectangle)
                        {
                            Color borderColor = FormatCell.BorderColor;
                            PictureStyle.Set("border-color", $"rgb({borderColor.R},{borderColor.G},{borderColor.B})");
                        }

                        Color bgColor = FormatCell.BgColor;
                        if (bgColor != Color.Empty)
                            PictureStyle.Set("background-color", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");

                    }


                    PictureStyle.Add("position", "absolute");
                    PictureStyle.Add("overflow", "hidden");

                    if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV))
                    {
                        PictureStyle.Remove("vertical-align");
                        if (FormatCell.bVertAlign == TextVertAlign.Middle)
                        {
                            PictureStyle.Add("display", "flex");
                            PictureStyle.Add("align-items", "center");
                        }
                    }

                    if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH))
                    {

                        if (FormatCell.bHorAlign.HasFlag(TextHorzAlign.BySelection) && FormatCell.bHorAlign.HasFlag(TextHorzAlign.Center))
                        {
                            PictureStyle.Set("text-align", "center");
                        }
                        else
                            PictureStyle.Set("text-align", FormatCell.bHorAlign.ToString());
                    }

                    CSSstyle LineStyle = new CSSstyle();
                    LineStyle.Set("stroke", "#000000");
                    FillLineStyle(FormatCell, ref LineStyle);
                    string SvgBackground = string.Empty;


                    if (DrawRectangle)
                    {
                        StringBuilder SVGPicture = new StringBuilder();
                        if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.PatternType))
                            LineStyle.Set("fill", "url(#defpattern)");
                        else
                            LineStyle.Set("fill", "none");

                        SVGPicture.AppendLine($"<svg baseProfile=\"full\" xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" height = \"{DrawingArea.Height + BorderWith}px\"  width = \"{DrawingArea.Width + BorderWith}px\" text-rendering=\"geometricPrecision\">");
                        SVGPicture.AppendLine(GetSVGFilPattern(FormatCell));
                        SVGPicture.AppendLine($"<g transform=\"translate({BorderWith / 2}, {BorderWith / 2 })\">");
                        SVGPicture.AppendLine($"<rect {LineStyle} x=\"1\" y=\"1\" width=\"{DrawingArea.Width}\" height=\"{DrawingArea.Height}\"/>");
                        SVGPicture.AppendLine("</g>");
                        SVGPicture.AppendLine("</svg>");
                        PictureStyle.Set("background-image", $"url(data:image/svg+xml;base64,{Convert.ToBase64String(Encoding.ASCII.GetBytes(SVGPicture.ToString()))})");
                    }

                    result.Write($"\t\t<div id=\"D{obj.Picture.dwZOrder}\"{PictureStyle}>\r\n");
                    switch (obj.Picture.dwType)
                    {
                        case ObjectType.Ole:
                        case ObjectType.Picture:
                            using (MemoryStream ms = new MemoryStream())
                            {
                                ///Странныый косяк с GDI+. Без такого финта выдает неопознанную ошибку
                                using (Bitmap bmp = obj.pObject)
                                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

                                result.Write($"\t\t\t<img src=\"data:image/png;base64,{Convert.ToBase64String(ms.ToArray())}\" width=\"{DrawingArea.Width + BorderWith}\" height=\"{DrawingArea.Height + BorderWith}\">\r\n");
                            }
                            break;
                        case ObjectType.Text:
                            CSSstyle TextStyle = new CSSstyle();
                            Text = FillTextStyle(obj, ref TextStyle);
                            TextStyle.Set("max-width", $"{DrawingArea.Width}px");
                            TextStyle.Set("width", $"{DrawingArea.Width}px");
                            TextStyle.Set("line-height", "1.57");
                            result.WriteLine($"<span{TextStyle}>{Text}</span>");
                            break;
                        case ObjectType.Line:
                            result.WriteLine($"<svg baseProfile=\"full\" xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" height = \"{Math.Max(DrawingArea.Height, 10)}px\"  width = \"{Math.Max(DrawingArea.Width, 10)}px\" text-rendering=\"geometricPrecision\">");
                            result.WriteLine("<g transform=\"translate(0.5, 0.5)\">");
                            Rectangle LineCoords = obj.ImageArea;

                            if (LineCoords.Height * LineCoords.Width >= 0)
                                result.WriteLine($"<line {LineStyle} x1=\"0\" y1=\"1\" x2=\"{DrawingArea.Width}\" y2=\"{Math.Max(DrawingArea.Height, BorderWith)}\"/>");
                            else
                                result.WriteLine($"<line {LineStyle} x1=\"0\" y2=\"1\" x2=\"{DrawingArea.Width}\" y1=\"{Math.Max(DrawingArea.Height, BorderWith)}\"/>");
                            result.WriteLine("</g>");
                            result.WriteLine("</svg>");
                            break;
                        default:
                            break;
                    }
                    result.Write("\t\t</div>\r\n");
                }

                result.Write("\t</body>\r\n");
                result.Write("</html>");
                //return result;
            }
        }


        private static string PictureBorderWidth(ObjectBorderWidth bPictureBorderWidth)
        {
            switch(bPictureBorderWidth)
            {
                case ObjectBorderWidth.Medium:
                    return "2px ";
                case ObjectBorderWidth.Thick:
                    return "3px ";
                case ObjectBorderWidth.Thin:
                default:
                    return "0.5px ";
            }
            
        }
        private static string PictureBorderStyle(CSheetFormat FormatCell)
        {
            if (FormatCell.bPictureBorderPresence == ObjectBorderPresence.Empty)
                return "none";

            switch (FormatCell.bPictureBorderStyle)
            {
                case ObjectBorderStyle.None:
                case ObjectBorderStyle.Solid:
                    return $"{PictureBorderWidth(FormatCell.bPictureBorderWidth)}solid";
                case ObjectBorderStyle.DashedShort:
                    return $"{PictureBorderWidth(FormatCell.bPictureBorderWidth)}dotted";
                case ObjectBorderStyle.DashedExtraLong:
                case ObjectBorderStyle.DashDotSparse:
                case ObjectBorderStyle.DashDotDot:
                    return $"{PictureBorderWidth(FormatCell.bPictureBorderWidth)}dashed";
                default:
                    return "none";
            }

        }
    }
}
