using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Drawing;
using System.IO.Compression;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Reflection;
using Ole;

namespace Moxel
{
     public class Moxel
    {
        const int UNITS_PER_INCH = 1440;
        const int UNITS_PER_PIXEL = UNITS_PER_INCH / 96;
        const int LOGPIXELSX = 88;
        const int LOGPIXELSY = 90;
        const int MOXEL_UNITS_PER_INCH = 288;
        const float CoordsCoeff = UNITS_PER_INCH / MOXEL_UNITS_PER_INCH;


        int Version = 6;
        public int nAllColumnCount; //всего колонок в таблице
        public int nAllRowCount;    //всего строк в таблице
        public int nAllObjectsCount;//Всего объектов

        Cellv6 DefFormat;
        Dictionary<int, LOGFONT> FontList;
        Dictionary<int, string> stringTable;
        DataCell TopColon;
        DataCell BottomColon;

        Dictionary<int, DataCell> Columns;
        Dictionary<int, MoxelRow> Rows;
        List<EmbeddedObject> Objects;
        List<CellsUnion> Unions;
        List<Section> VerticalSections;
        List<Section> HorisontalSections;
        int[] HorisontalPageBreaks;
        int[] VerticalPageBreaks;
        List<MoxelArea> AreaNames;

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

        string GetBorderStyle(BorderStyle Moxelborder, int borderColorIndex)
        {
            uint borderColor = 0;

            if (borderColorIndex >= 0 || borderColorIndex < a1CPallete.Length)
                borderColor = a1CPallete[borderColorIndex];

            switch(Moxelborder)
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


        public int GetColumnWidth(int ColNumber)
        {
            int result = 0;
            if (DefFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                result = DefFormat.wWidth;

            if(Columns.ContainsKey(ColNumber))
                if(Columns[ColNumber].FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                    result = Columns[ColNumber].FormatCell.wWidth;

            if (result == 0)
                result = 40;

            return result;
        }

        public int GetWidth(int x1, int x2)
        {
            int width = 0;
            for (int i = x1; i < x2; i++)
                width += GetColumnWidth(i);

            return width;
        }

        public int GetRowHeight(int RowNumber)
        {
            int result = 0;
            if (DefFormat.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                result = DefFormat.wHeight;

            if (Rows.ContainsKey(RowNumber))
                if (Rows[RowNumber].FormatCell.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                    result = Rows[RowNumber].FormatCell.wHeight;

            if (result == 0)
                result = 45;

            return result;
        }

        private int GetHeight(int y1, int y2)
        {
            int height = 0;
            for (int i = y1; i < y2; i++)
                height += GetRowHeight(i);

            return height;
        }

        public int GetSpan(int RowNumber, int ColumnNumber)
        {
            MoxelRow Row = null; 
            if (Rows.ContainsKey(RowNumber))
                Row = Rows[RowNumber];
            else
                return nAllColumnCount - ColumnNumber;
            int i = ColumnNumber;
            for (i = ColumnNumber + 1; i < nAllColumnCount; i++)
            {
                if (Row.Cells.ContainsKey(i))
                    return i - ColumnNumber - 1;
                
            }

            return 0;
        }

        public void FillFormat(Cellv6 FormatCell, ref CSSstyle CellStyle, ref string Text)
        {
            string FontFamily = string.Empty;
            float FontSize = 0;
            FillFormat(FormatCell, ref CellStyle, ref Text, ref FontFamily, ref FontSize);
        }

        public void FillFormat(Cellv6 FormatCell, ref CSSstyle CellStyle, ref string Text, ref string FontFamily, ref float FontSize)
        {
            if (!string.IsNullOrWhiteSpace(Text))
            {
                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontName))
                {
                    FontFamily = FontList[FormatCell.wFontNumber].lfFaceName;
                    CellStyle.Set("font-family", FontFamily);
                }

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                {
                    FontSize = (float)-FormatCell.wFontSize / 4;
                    CellStyle.Set("font-size", $"{Math.Round(FontSize, 0)}pt");
                }

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontWeight))
                    CellStyle.Set("font-weight", FormatCell.bFontBold.ToString());

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontItalic))
                    if(FormatCell.bFontItalic)
                        CellStyle.Set("font-style", "italic");

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontColor))
                {
                    int ColorIndex = FormatCell.bFontColor;
                    if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                    {
                        Color bgColor = Color.FromArgb((int)(a1CPallete[ColorIndex] + 0xFF000000));
                        CellStyle.Set("color", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");
                    }
                }

                if (FormatCell.bControlContent != TextControl.Wrap)
                {
                    Text = Text.Replace(" ", "&nbsp;");
                    CellStyle.Set("white-space", "nowrap");
                    CellStyle.Set("max-width", "0px");
                }

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV))
                    CellStyle.Set("vertical-align", FormatCell.bVertAlign.ToString());

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH))
                    CellStyle.Set("text-align", FormatCell.bHorAlign.ToString());
            }

            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
            {
                int ColorIndex = FormatCell.bBackground;
                if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                {
                    Color bgColor = Color.FromArgb((int)(a1CPallete[ColorIndex] + 0xFF000000));
                    CellStyle.Set("background-color", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");
                }
            }
        }

        public void FillLineStyle(Cellv6 FormatCell, ref CSSstyle LineStyle)
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

            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderColor))
            {
                int ColorIndex = FormatCell.bBorderColor;
                if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                {
                    Color bgColor = Color.FromArgb((int)(a1CPallete[ColorIndex] + 0xFF000000));
                    LineStyle.Set("stroke", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");
                }
            }
        }


        public bool SaveToHtml(string filename)
        {
            Cellv6 FormatCell = DefFormat;
            string DefFontNAme = string.Empty;
            float DefFontSize = 8.0f;

            if (FontList.Count == 1)
                DefFontNAme = $" style=\"font-family:{FontList.First().Value.lfFaceName}\"";

            if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                DefFontSize = -(float)FormatCell.wFontSize / 4;





            StringBuilder result = new StringBuilder("<!DOCTYPE HTML PUBLIC \" -//W3C//DTD HTML 5.0 Transitional//EN\">\r\n<HTML>\r\n");
            result.AppendLine("<HEAD>\r\n<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html; CHARSET = utf-8\"/>");

            result.AppendLine("<style type=\"text/css\">");
            result.AppendLine("body { background: #ffffff; margin: 0; font-family: Arial; font-size: 8pt; font-style: normal; }");
            result.AppendLine("table {table-layout: fixed; padding: 0px; padding-left: 2px; vertical-align:bottom; border-collapse:collapse;width: 100%; font-family: Arial; font-size: 8pt; font-style: normal; }");
            result.AppendLine("td { padding: 0px 0px 0px 2px;}");
            result.AppendLine("tr { height: 15px;}");
            result.AppendLine("</style>");

            result.AppendLine("</HEAD>");
            result.AppendLine($"\t<body{DefFontNAme}\">");
            result.AppendLine($"\t\t<TABLE style=\"width: {Math.Round(GetWidth(0, nAllColumnCount) * 0.875) }px; height: 0px; \" border=0 CELLSPACING=0>");
            result.AppendLine($"\t\t\t<colgroup>");
            for (int columnnumber = 0; columnnumber < nAllColumnCount; columnnumber++)
            {
                string Width = " width=35";

                if (Columns.ContainsKey(columnnumber))
                    FormatCell = Columns[columnnumber];
                else
                    FormatCell = DefFormat;

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                    Width = $" width={Math.Round(FormatCell.wWidth * 0.875).ToString()}";


                result.AppendLine($"\t\t\t\t<col{Width}/>");
            }
            result.Append($"\t\t\t</colgroup>\r\n");
            result.Append($"\t\t\t<tbody>\r\n");

            for (int rownumber = 0; rownumber < nAllRowCount; rownumber++)
            {
                MoxelRow Row = null;
                FormatCell = DefFormat;
                CSSstyle RowStyle = new CSSstyle();
                StringBuilder RowString = new StringBuilder();

                if (Rows.ContainsKey(rownumber))
                {
                    Row = Rows[rownumber];
                    FormatCell = Row.FormatCell;
                }
                
                List<CellsUnion> RowUnion = Unions.Where(t => (t.dwTop <= rownumber && t.dwBottom >= rownumber)).ToList();

                for (int columnnumber = 0; columnnumber < nAllColumnCount; columnnumber++)
                {
                    string Text = string.Empty;
                    CSSstyle CellStyle = new CSSstyle();

                    CellsUnion Union = RowUnion.FirstOrDefault(t => (t.dwLeft <= columnnumber && t.dwRight >= columnnumber));

                    if (Row != null)
                    {
                        if (Row.Cells.ContainsKey(columnnumber))
                        {
                            Text = Row.Cells[columnnumber].Text;
                            FormatCell = Row.Cells[columnnumber].FormatCell;
                            if(Row.Cells[columnnumber].TextOrientation != 0)
                            CellStyle.Set("transform",$"rotate(-{Row.Cells[columnnumber].TextOrientation}deg)");
                        }
                    }
                    string FontFamily = DefFontNAme;
                    float FontSize = DefFontSize;

                    FillFormat(FormatCell, ref CellStyle, ref Text, ref FontFamily, ref FontSize);

                    if (!string.IsNullOrWhiteSpace(Text))
                    {
                        Text = Text.Replace("\r\n", "<br>");
                        
                        Size Constr = new Size { Width = (int)Math.Round((GetWidth(columnnumber, columnnumber + Union.ColumnSpan) + GetColumnWidth(columnnumber)) * 0.875), Height = 0 };
                        Size textsize = System.Windows.Forms.TextRenderer.MeasureText(Text.TrimStart(' '), new Font(FontFamily, FontSize, FormatCell.bFontBold == clFontWeight.Bold ? FontStyle.Bold : FontStyle.Regular), Constr, TextFormatFlags.WordBreak | TextFormatFlags.ExternalLeading);

                        textsize.Height /= Union.RowSpan + 1;

                        if (!Row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                            Row.FormatCell.dwFlags = Row.FormatCell.dwFlags | MoxelCellFlags.RowHeight;

                        Row.FormatCell.wHeight = Math.Max((short)(Math.Max(textsize.Height, 15) * 3), Row.FormatCell.wHeight);
                        
                    }

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

                    string span = string.Empty;
                    int colspan = GetSpan(rownumber, columnnumber);

                    if (Union.ColumnSpan > 0)
                    {
                        span = Union.HtmlSpan;
                        colspan = Union.ColumnSpan;
                    }
                    else
                        if (colspan > 0)
                    {
                        span = $" colspan=\"{colspan}\"";
                    }


                    if (!Union.ContainsCell(rownumber, columnnumber))
                        if(!string.IsNullOrWhiteSpace(Text))
                            RowString.AppendLine($"\t\t\t\t<td{span}{CellStyle}>{Text}</td>");
                        else
                            RowString.AppendLine($"\t\t\t\t<td{span}{CellStyle}/>");

                    columnnumber += colspan;
                }

                if (Row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                {
                    RowStyle.Set("height", $"{Row.FormatCell.wHeight / 3}px");
                }

                result.AppendLine($"\t\t\t<tr{RowStyle} id=\"R{rownumber:00}\">\r\n{RowString}\r\n\t\t\t</tr>");
                Rows[rownumber] = Row;
            }
            result.AppendLine($"\t\t\t</tbody>");
            result.AppendLine("\t\t</table>");

            foreach(EmbeddedObject obj in Objects)
            {
                CSSstyle PictureStyle = new CSSstyle();
                CSSstyle TextStyle = new CSSstyle();
                Rectangle DrawingArea = obj.AbsoluteImageArea;

                FormatCell = obj.Format;

                int BorderWith = (int)FormatCell.bPictureBorderWidth * 2 + 1;


                PictureStyle.Add("top", $"{DrawingArea.Top - BorderWith}px");
                PictureStyle.Add("left", $"{DrawingArea.Left - BorderWith}px");

                PictureStyle.Add("width", $"{DrawingArea.Width - BorderWith}px");
                PictureStyle.Add("height", $"{DrawingArea.Height - BorderWith}px");

                if (obj.Picture.dwType != ObjectType.Line)
                {
                    if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                    {
                        int ColorIndex = FormatCell.bBackground;
                        if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                        {
                            Color bgColor = Color.FromArgb((int)(a1CPallete[ColorIndex] + 0xFF000000));
                            PictureStyle.Set("background-color", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");
                        }
                    }


                    if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight) || FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
                    {
                        if ((FormatCell.bPictureBorderPresence != ObjectBorderPresence.All) && FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                        {
                            if (FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Left))
                                PictureStyle.Set("border-left", FormatCell.bPictureBorderStyle.ToString());
                            if (FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Right))
                                PictureStyle.Set("border-right", FormatCell.bPictureBorderStyle.ToString());
                            if (FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Top))
                                PictureStyle.Set("border-top", FormatCell.bPictureBorderStyle.ToString());
                            if (FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Bottom))
                                PictureStyle.Set("border-bottom", FormatCell.bPictureBorderStyle.ToString());
                        }
                        else
                            PictureStyle.Set("border", FormatCell.bPictureBorderStyle.ToString());

                        PictureStyle.Set("border-width", $"{(byte)FormatCell.bPictureBorderWidth * 2 + 1}px");

                        PictureStyle.Set("border-color", "#000000");
                    }
                    else
                    {
                        if(!FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
                            PictureStyle.Set("border", "#000000 solid 1px");
                        else
                            PictureStyle.Set("border", $"#000000 solid {(byte)FormatCell.bPictureBorderWidth * 2 + 1}px");
                    }

                    if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderColor))
                    {
                        int ColorIndex = FormatCell.bBorderColor;
                        if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                        {
                            Color bgColor = Color.FromArgb((int)(a1CPallete[ColorIndex] + 0xFF000000));
                            PictureStyle.Set("border-color", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");
                        }
                    }

                    FillFormat(obj.Format, ref TextStyle, ref obj.Format.Text);
                    TextStyle.Set("max-width", $"{DrawingArea.Width}px");
                    TextStyle.Set("width", $"{DrawingArea.Width}px");
                    TextStyle.Set("line-height", "1.57");
                    if (obj.Format.TextOrientation != 0)
                        TextStyle.Set("transform", $"rotate(-{obj.Format.TextOrientation}deg)");
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

                CSSstyle LineStyle = new CSSstyle();
                LineStyle.Set("stroke", "#000000");
                FillLineStyle(FormatCell, ref LineStyle);



                result.Append($"\t\t<div id=\"D{obj.Picture.dwZOrder}\"{PictureStyle}>\r\n");
                string DataUri = string.Empty;
                switch (obj.Picture.dwType)
                {
                    case ObjectType.Ole:
                    case ObjectType.Picture:
                            using (MemoryStream ms = new MemoryStream())
                            {
                                ((Bitmap)obj.pObject)?.Save(ms, ImageFormat.Png);
                                DataUri = Convert.ToBase64String(ms.ToArray());
                                result.Append($"\t\t\t<img src=\"data:image/png;base64,{DataUri}\" width=\"{DrawingArea.Width + BorderWith}\" height=\"{DrawingArea.Height + BorderWith}\">\r\n");
                            }
                        break;
                    case ObjectType.Text:
                        result.AppendLine($"<span{TextStyle}>{obj.Format.Text.Replace("\r\n", "<br>")}</span>");
                        break;
                    case ObjectType.Line:
                        result.AppendLine($"<svg baseProfile=\"full\" xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" height = \"{Math.Max(DrawingArea.Height,10)}px\"  width = \"{Math.Max(DrawingArea.Width,10)}px\" text-rendering=\"geometricPrecision\">");
                        result.AppendLine("<g transform=\"translate(0.5, 0.5)\">");
                        Rectangle LineCoords = obj.ImageArea;
                        
                        if (LineCoords.Height * LineCoords.Width >=0)
                            result.AppendLine($"<line {LineStyle} x1=\"0\" y1=\"1\" x2=\"{DrawingArea.Width}\" y2=\"{Math.Max(DrawingArea.Height, BorderWith)}\"/>");
                        else
                            result.AppendLine($"<line {LineStyle} x1=\"0\" y2=\"1\" x2=\"{DrawingArea.Width}\" y1=\"{Math.Max(DrawingArea.Height, BorderWith)}\"/>");
                        result.AppendLine("</g>");
                        result.AppendLine("</svg>");
                        break;
                    case ObjectType.Rectangle:

                        if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.PatternType))
                            LineStyle.Set("fill", "url(#defpattern)");
                        else
                            LineStyle.Set("fill", "none");

                        result.AppendLine($"<svg baseProfile=\"full\" xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" height = \"{DrawingArea.Height + BorderWith}px\"  width = \"{DrawingArea.Width + BorderWith}px\" text-rendering=\"geometricPrecision\">");
                        result.AppendLine(GetSVGFilPattern(FormatCell));
                        result.AppendLine("<g transform=\"translate(0.5, 0.5)\">");
                        result.AppendLine($"<rect {LineStyle} x=\"{BorderWith / 2}\" y=\"{BorderWith / 2}\" width=\"{DrawingArea.Width}\" height=\"{DrawingArea.Height}\"/>");
                        result.AppendLine("</g>");
                        result.AppendLine("</svg>");
                        break;
                    default:
                        break;
                }
                result.Append("\t\t</div>\r\n");
            }

            result.Append("\t</body>\r\n");
            result.Append("</html>");
            File.WriteAllText(filename, result.ToString(), Encoding.UTF8);
            return true;
        }

        private string GetSVGFilPattern(Cellv6 formatCell)
        {
            Color PatternColor = Color.Black;
            if (formatCell.dwFlags.HasFlag(MoxelCellFlags.PatternColor))
            {
                int ColorIndex = (int)formatCell.bPatternColor;
                if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                {
                    PatternColor = Color.FromArgb((int)(a1CPallete[ColorIndex]));
                }
            }
            StringBuilder result = new StringBuilder($"<style type=\"text/css\">\r\n\t#defpattern {{ fill: rgb({PatternColor.R},{PatternColor.G}, {PatternColor.G}); }}\r\n\t</style>");

            if (formatCell.dwFlags.HasFlag(MoxelCellFlags.PatternType))
            {
                switch (formatCell.bPatternType)
                {
                    case 1:
                        result.AppendLine("<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"#FFFFFF\"/>\r\n\t\t<rect x=\"0\" y=\"1\" width=\"2\" height=\"1\"></rect>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    case 2:
                        result.AppendLine("<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"#FFFFFF\"/>\r\n\t\t<rect x=\"1\" y=\"1\" width=\"1\" height=\"1\"></rect>\r\n\t\t\t<rect x=\"0\" y=\"1\" width=\"1\" height=\"1\" fill=\"#FFFFFF\"/>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    case 3:
                        result.AppendLine("<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"2\" height=\"2\">\r\n\t\t<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\"/>\r\n\t\t\t<rect x=\"1\" y=\"0\" width=\"1\" height=\"1\" fill=\"#FFFFFF\"/>\r\n\t\t<rect x=\"0\" y=\"1\" width=\"2\" height=\"1\" fill=\"#FFFFFF\"/>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    case 4:
                        result.AppendLine("<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"8\" height=\"8\"><rect x=\"0\" y=\"0\" width=\"8\" height=\"8\"/>\r\n\t\t<rect x=\"0\" y=\"0\" width=\"3\" height=\"2\" fill=\"#FFFFFF\"/>\r\n\t\t<rect x=\"4\" y=\"0\" width=\"4\" height=\"2\" fill=\"#FFFFFF\"/>\r\n\t\t<rect x=\"0\" y=\"3\" width=\"7\" height=\"3\" fill=\"#FFFFFF\"/>\r\n\t\t<rect x=\"0\" y=\"7\" width=\"3\" height=\"1\" fill=\"#FFFFFF\"/>\r\n\t\t<rect x=\"4\" y=\"7\" width=\"4\" height=\"1\" fill=\"#FFFFFF\"/>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    case 5:
                        result.AppendLine($"<defs>\r\n\t<pattern id=\"defpattern\" patternUnits=\"userSpaceOnUse\" width=\"8\" height=\"8\"><rect stroke-width=\"0\" height=\"8\" width=\"8\" y=\"0\" x=\"0\" fill=\"#FFFFFF\"></rect><path d=\" M 0 0.5 H 1 M 4 0.5 H 5 M 1 1.5 H 2 M 3 1.5 H 4 M 5 1.5 H 6 M 2 2.5 H 3 M 6 2.5 H 7 M 1 3.5 H 2 M 5 3.5 H 6 M 7 3.5 H 8 M 0 4.5 H 1 M 4 4.5 H 5 M 3 5.5 H 4 M 5 5.5 H 6 M 7 5.5 H 8 M 2 6.5 H 3 M 6 6.5 H 7 M 1 7.5 H 2 M 3 7.5 H 4 M 7 7.5 H 8\" stroke-width=\"1\" stroke=\"rgb({PatternColor.R},{PatternColor.G}, {PatternColor.G})\"/>\r\n\t</pattern>\r\n\t</defs>");
                        break;
                    default:
                        break;
                }
            }
            else
                return string.Empty;

            return result.ToString();
        }



        #region Перечисления
        [Flags]
        public enum MoxelCellFlags : uint
        {
            Empty = 0x00000000,
            FontName = 0x00000001,
            FontSize = 0x00000002,
            FontWeight = 0x00000004,
            FontItalic = 0x00000008,
            FontUnderline = 0x00000010,
            /// <summary>
            /// Так же PictureBorderStyle 
            /// </summary>
            BorderLeft = 0x00000020,   //   PictureBorderStyle = 0x00000020,
            /// <summary>
            /// Так же PictureBorderWidth
            /// </summary>
            BorderTop = 0x00000040,    //   PictureBorderWidth = 0x00000040,
            /// <summary>
            /// Так же PictureBorderPresence
            /// </summary>
            BorderRight = 0x00000080,  //   PictureBorderPresence = 0x00000080,
            BorderBottom = 0x00000100, //   PicturePrint = 0x00000100,
            BorderColor = 0x00000200,
            /// <summary>
            /// Так же ColumnPagePosition
            /// </summary>
            RowHeight = 0x00000400,    //   ColumnPagePosition = 0x00000400,
            ColumnWidth = 0x00000800,
            AlignH = 0x00001000,
            AlignV = 0x00002000,
            FontColor = 0x00004000,
            Background = 0x00008000,
            PatternType = 0x00010000,
            PatternColor = 0x00020000,
            Control = 0x00040000,
            Type = 0x00080000,
            Protect = 0x00100000,
            Data = 0x00200000,
            TextOrientation = 0x00400000,
            Value = 0x40000000,
            Text = 0x80000000
        };



        public enum TextHorzAlign : byte
        {
            Left = 0,
            Right = 2,
            Justify = 4,
            Center = 6,
            BySelection = 0x20
        };

        //UINT const AlignMask = 0x07;
        public enum TextVertAlign : byte
        {
            Top = 0,
            Bottom = 8,
            Middle = 0x18
        };

    public enum FontWeight : int
        {
            FW_DONTCARE = 0,
            FW_THIN = 100,
            FW_EXTRALIGHT = 200,
            FW_LIGHT = 300,
            FW_NORMAL = 400,
            FW_MEDIUM = 500,
            FW_SEMIBOLD = 600,
            FW_BOLD = 700,
            FW_EXTRABOLD = 800,
            FW_HEAVY = 900,
        }
        public enum FontCharSet : byte
        {
            ANSI_CHARSET = 0,
            DEFAULT_CHARSET = 1,
            SYMBOL_CHARSET = 2,
            SHIFTJIS_CHARSET = 128,
            HANGEUL_CHARSET = 129,
            HANGUL_CHARSET = 129,
            GB2312_CHARSET = 134,
            CHINESEBIG5_CHARSET = 136,
            OEM_CHARSET = 255,
            JOHAB_CHARSET = 130,
            HEBREW_CHARSET = 177,
            ARABIC_CHARSET = 178,
            GREEK_CHARSET = 161,
            TURKISH_CHARSET = 162,
            VIETNAMESE_CHARSET = 163,
            THAI_CHARSET = 222,
            EASTEUROPE_CHARSET = 238,
            RUSSIAN_CHARSET = 204,
            MAC_CHARSET = 77,
            BALTIC_CHARSET = 186,
        }
        public enum FontPrecision : byte
        {
            OUT_DEFAULT_PRECIS = 0,
            OUT_STRING_PRECIS = 1,
            OUT_CHARACTER_PRECIS = 2,
            OUT_STROKE_PRECIS = 3,
            OUT_TT_PRECIS = 4,
            OUT_DEVICE_PRECIS = 5,
            OUT_RASTER_PRECIS = 6,
            OUT_TT_ONLY_PRECIS = 7,
            OUT_OUTLINE_PRECIS = 8,
            OUT_SCREEN_OUTLINE_PRECIS = 9,
            OUT_PS_ONLY_PRECIS = 10,
        }

        [Flags]
        public enum FontClipPrecision : byte
        {
            CLIP_DEFAULT_PRECIS = 0,
            CLIP_CHARACTER_PRECIS = 1,
            CLIP_STROKE_PRECIS = 2,
            CLIP_MASK = 0xf,
            CLIP_LH_ANGLES = (1 << 4),
            CLIP_TT_ALWAYS = (2 << 4),
            CLIP_DFA_DISABLE = (4 << 4),
            CLIP_EMBEDDED = (8 << 4),
        }
        public enum FontQuality : byte
        {
            DEFAULT_QUALITY = 0,
            DRAFT_QUALITY = 1,
            PROOF_QUALITY = 2,
            NONANTIALIASED_QUALITY = 3,
            ANTIALIASED_QUALITY = 4,
            CLEARTYPE_QUALITY = 5,
            CLEARTYPE_NATURAL_QUALITY = 6,
        }

        [Flags]
        public enum FontPitchAndFamily : byte
        {
            DEFAULT_PITCH = 0,
            FIXED_PITCH = 1,
            VARIABLE_PITCH = 2,
            FF_DONTCARE = (0 << 4),
            FF_ROMAN = (1 << 4),
            FF_SWISS = (2 << 4),
            FF_MODERN = (3 << 4),
            FF_SCRIPT = (4 << 4),
            FF_DECORATIVE = (5 << 4),
        }

        public enum TextControl : byte
        {
            Auto = 0,
            Cut = 1,
            Fill = 2,
            Wrap = 3,
            Red = 4,
            FillAndRed = 5
        };
        #endregion

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        public struct LOGFONT
        {
            public int lfHeight;
            public int lfWidth;
            public int lfEscapement;
            public int lfOrientation;
            public FontWeight lfWeight;
            [MarshalAs(UnmanagedType.U1)]
            public bool lfItalic;
            [MarshalAs(UnmanagedType.U1)]
            public bool lfUnderline;
            [MarshalAs(UnmanagedType.U1)]
            public bool lfStrikeOut;
            public FontCharSet lfCharSet;
            public FontPrecision lfOutPrecision;
            public FontClipPrecision lfClipPrecision;
            public FontQuality lfQuality;
            public FontPitchAndFamily lfPitchAndFamily;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string lfFaceName;

            public override string ToString()
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("lfCharSet: {0}\n", lfCharSet);
                sb.AppendFormat("lfFaceName: {0}\n", lfFaceName);

                return sb.ToString();
            }
        }

        [StructLayout(LayoutKind.Explicit, CharSet = CharSet.Ansi, Pack = 1)]
        public struct Cellv6
        {
            [FieldOffset(0x00)] public MoxelCellFlags dwFlags; // MoxcelCellFlags
        // union{
            [FieldOffset(0x04)] public short wShow; // 1 - да, 0xFFFF - нет. Используется в колонтитулах
            [FieldOffset(0x04)] public short wColumnPosition; // Используется в колонках
            [FieldOffset(0x04)] public short wHeight; // Используется в строках
        //}
        // union{
            [FieldOffset(0x06)] public short wStartPage; // Колонтитулы
            [FieldOffset(0x06)] public short wWidth; // Колонки
            [FieldOffset(0x06)] public short wRowPosition; // Строки
        //}
            [FieldOffset(0x08)] public short wFontNumber;
            [FieldOffset(0x0A)] public short wFontSize;
            [FieldOffset(0x0C)] public clFontWeight bFontBold;
            [MarshalAs(UnmanagedType.U1)]
            [FieldOffset(0x0D)] public bool bFontItalic;
            [MarshalAs(UnmanagedType.U1)]
            [FieldOffset(0x0E)] public bool bFontUnderline;
            [FieldOffset(0x0F)] public TextHorzAlign bHorAlign;
            [FieldOffset(0x10)] public TextVertAlign bVertAlign;
            [FieldOffset(0x11)] public byte bPatternType;
        // union {
            [FieldOffset(0x12)] public BorderStyle bBorderLeft;
            [FieldOffset(0x12)] public ObjectBorderStyle bPictureBorderStyle;
                                                                             //};
                                                                             // union {
            [FieldOffset(0x13)] public BorderStyle bBorderTop;
            [FieldOffset(0x13)] public ObjectBorderWidth bPictureBorderWidth;
        //};
        // union {
            [FieldOffset(0x14)] public BorderStyle bBorderRight;
            [FieldOffset(0x14)] public ObjectBorderPresence bPictureBorderPresence;
            //};
            // union {
            [MarshalAs(UnmanagedType.U1)]
            [FieldOffset(0x15)] public bool bPrintPicture;
            [FieldOffset(0x15)] public BorderStyle bBorderBottom;
        //};
            [FieldOffset(0x16)] public byte bPatternColor;
            [FieldOffset(0x17)] public byte bBorderColor;
            [FieldOffset(0x18)] public byte bFontColor;
            [FieldOffset(0x19)] public byte bBackground;
            [FieldOffset(0x1A)] public TextControl bControlContent; // MoxcelControlContent
            [FieldOffset(0x1B)] public ContentType bType; // MoxcelContentType
            [MarshalAs(UnmanagedType.U1)]
            [FieldOffset(0x1C)] public bool bAllowEdit;
            [FieldOffset(0x1D)] public byte bXZ1;

            public static Cellv6 Empty = new Cellv6
            {
                dwFlags = MoxelCellFlags.Empty,
                wShow = 0,
                wColumnPosition = 0,
                wHeight = 0,
                wStartPage = 0,
                wWidth = 0,
                wRowPosition = 0,
                wFontNumber = 0,
                wFontSize = 0,
                bFontBold = clFontWeight.Normal,
                bFontItalic = false,
                bFontUnderline = false,
                bHorAlign = 0,
                bVertAlign = 0,
                bPatternType = 0,
                bBorderLeft = BorderStyle.None,
                bPictureBorderStyle = ObjectBorderStyle.None,
                bBorderTop = BorderStyle.None,
                bPictureBorderWidth = ObjectBorderWidth.Thin,
                bBorderRight = BorderStyle.None,
                bPictureBorderPresence = ObjectBorderPresence.All,
                bBorderBottom = BorderStyle.None,
                bPrintPicture = true,
                bPatternColor = 0,
                bBorderColor = 0,
                bFontColor = 0,
                bBackground = 0,
                bControlContent = 0,
                bType = ContentType.Text,
                bAllowEdit = true,
                bXZ1 = 0
            };
        }

        public class DataCell
        {
            public Cellv6 FormatCell;
            public string Text;
            public string Value;
            public byte[] Data;
            public short TextOrientation = 0;
            Moxel Parent;

            public DataCell()
            { }
            public DataCell(BinaryReader br, Moxel parent = null)
            {
                Parent = parent;
                FormatCell = br.Read<Cellv6>();

                if (Parent != null)
                    if (Parent.Version == 7)
                        TextOrientation = br.ReadInt16();

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Text))
                {
                    Text = br.ReadCString();
                }

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Value))
                {
                    Value = br.ReadCString();
                }

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Data))
                {
                    Data = br.ReadBytes(br.ReadCount());
                }
            }

            public static implicit operator Cellv6(DataCell dc)
            {
                return dc.FormatCell;
            }

        }
         
        public enum ContentType : byte
        {
            Text = 0,
            Expression = 1,
            Pattern = 2,
            FixedPattern = 3
        };

        public enum ObjectType 
        {
            Line = 1,   //1-линия
            Rectangle,  //2-квадрат
            Text,       //3-блок текста (но без текста)
            Ole,        //4-ОЛЕ обьект (в т.ч. диаграмма 1С)
            Picture     //5-картинка
        };

        const short  BMPSignature = 0x4D42; // "BM"
        const uint WMFSignature = 0x9AC6CDD7; // placeable WMF

        public class EmbeddedObject
        {

            public enum PictureType
            {
                Bitmap,
                WMF,
                EMF,
                Unknown
            }

            public Rectangle AbsoluteImageArea { get
                {
                    int left=0, top=0, right=0, bottom=0;

                        left = (int)Math.Round(Parent.GetWidth(0, Picture.dwColumnStart) * 0.875 + Picture.dwOffsetLeft / 2.8, 0);
                        right = (int)Math.Round(Parent.GetWidth(0, Picture.dwColumnEnd) * 0.875 + Picture.dwOffsetRight / 2.8, 0);

                        top = (int)Math.Round(Parent.GetHeight(0, Picture.dwRowStart) / 3 + Picture.dwOffsetTop / 2.8, 0);
                        bottom = (int)Math.Round(Parent.GetHeight(0, Picture.dwRowEnd) / 3 + Picture.dwOffsetBottom / 2.8, 0);

                    return new Rectangle { X = Math.Min(left, right) , Y = Math.Min(top, bottom), Width = Math.Abs(right - left), Height = Math.Abs(bottom - top)};
                }
            }

            public Rectangle ImageArea
            {
                get
                {
                    int left = 0, top = 0, right = 0, bottom = 0;

                    left = (int)Math.Round(Parent.GetWidth(0, Picture.dwColumnStart) * 0.875 + Picture.dwOffsetLeft / 2.8, 0);
                    right = (int)Math.Round(Parent.GetWidth(0, Picture.dwColumnEnd) * 0.875 + Picture.dwOffsetRight / 2.8, 0);

                    top = (int)Math.Round(Parent.GetHeight(0, Picture.dwRowStart) / 3 + Picture.dwOffsetTop / 2.8, 0);
                    bottom = (int)Math.Round(Parent.GetHeight(0, Picture.dwRowEnd) / 3 + Picture.dwOffsetBottom / 2.8, 0);

                    return new Rectangle { X = left, Y = top, Width = right - left, Height = bottom - top};
                }
            }

            Moxel Parent = null;
            public DataCell Format;
            public Picture Picture;
            public object pObject;
            public object OleObject;
            public string ProgID;
            public Guid ClsId;
            public byte[] OleObjectStorage;

            public EmbeddedObject()
            { }
            public EmbeddedObject(BinaryReader br, Moxel parent)
            {
                Parent = parent;
                Format = br.Read<DataCell>(parent);
                Picture = br.Read<Picture>();
                switch (Picture.dwType)
                {
                    case ObjectType.Picture:
                        pObject = LoadPicture(Picture, br );
                        break;
                    case ObjectType.Line:
                    case ObjectType.Rectangle:
                    case ObjectType.Text:
                        break;
                    case ObjectType.Ole:
                        pObject = LoadOleObject(br);
                        break;
                    default:
                        throw new Exception("Неизвестный тип внедренного объекта");
                }
            }

            [DllImport("gdi32.dll",  SetLastError = true)]
            static extern int GetDeviceCaps(IntPtr hDC, int Caps);

            object LoadOleObject(BinaryReader br)
            {
                string classname = string.Empty;
                short wClassNameFlag = br.ReadInt16();
                if (wClassNameFlag == -1)
                {
                    br.ReadInt16();
                    short wClassNameLength = br.ReadInt16();
                    classname = Encoding.GetEncoding(1251).GetString(br.ReadBytes(wClassNameLength));
                }

                int dwObjectType = br.ReadInt32();
                int dwItemNumber = br.ReadInt32();
                int dwAspect = br.ReadInt32();
                short wUseMoniker = br.ReadInt16();

                dwAspect = br.ReadInt32();

                int dwObjectSize = br.ReadInt32();
                OleObjectStorage = br.ReadBytes(dwObjectSize);

                OLE32.CoInitializeEx(IntPtr.Zero, OLE32.CoInit.ApartmentThreaded); //COINIT_APARTMENTTHREADED
                OLE32.ILockBytes LockBytes;
                OLE32.IStorage RootStorage;

                IntPtr hGlobal = Marshal.AllocHGlobal(OleObjectStorage.Length);
                Marshal.Copy(OleObjectStorage, 0, hGlobal, OleObjectStorage.Length);
                OLE32.CreateILockBytesOnHGlobal(hGlobal, false, out LockBytes);
                OLE32.IOleObject pOle = null;

                HRESULT result = OLE32.StgOpenStorageOnILockBytes(LockBytes, null, OLE32.STGM.STGM_READWRITE | OLE32.STGM.STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out RootStorage);
                
                System.Runtime.InteropServices.ComTypes.STATSTG MetaDataInfo;
                RootStorage.Stat(out MetaDataInfo, 0);
                ClsId = MetaDataInfo.clsid;
                OLE32.ProgIDFromCLSID(ref ClsId, out ProgID);

                Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

                IOleClientSite ole_cs = null;
                result = OLE32.OleLoad(RootStorage, ref IID_IUnknown, ole_cs, out pOle);
                if (result != HRESULT.S_OK)
                {
                    int res = Marshal.GetLastWin32Error();
                    return null;
                }

                IntPtr pUnknwn = Marshal.GetIUnknownForObject(pOle);
                result = OLE32.OleSetContainedObject(pUnknwn, true);
                result = OLE32.OleNoteObjectVisible(pUnknwn, true);
                Marshal.Release(pUnknwn);
                result = OLE32.OleRun(pOle);

                //tagSIZEL sizel = new tagSIZEL();
                //pOle.GetExtent(1, ref sizel);
                //float LogUnitsPerDevPixel_X = 1, LogUnitsPerDevPixel_Y = 1;

                //using (Graphics g = Graphics.FromImage(new Bitmap(1, 1)))
                //{
                //    HandleRef hdcsrc = new HandleRef(g, g.GetHdc());
                //    LogUnitsPerDevPixel_X = ((float)UNITS_PER_INCH) / ((float)GetDeviceCaps(hdcsrc.Handle, LOGPIXELSX));
                //    LogUnitsPerDevPixel_Y = ((float)UNITS_PER_INCH) / ((float)GetDeviceCaps(hdcsrc.Handle, LOGPIXELSY));
                //    g.ReleaseHdc(hdcsrc.Handle);
                //}

                Rectangle Size = AbsoluteImageArea;

                Rectangle rect = new Rectangle(0, 0, Size.Width * 3, Size.Height * 3 );
                Bitmap m = new Bitmap(rect.Width, rect.Height);

                using (Graphics g = Graphics.FromImage(m))
                {
                    Color bgColor = Color.White;

                    
                    bool MakeTransparent = false;
                    if (ProgID == "BMP1C.Bmp1cCtrl.1")
                        if ((pOle as _DBmp_1c).GrMode == 1) // Иначе рисует только маску
                        {
                            (pOle as _DBmp_1c).GrMode = 2;
                            MakeTransparent = true;
                        }

                    if (Format.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                    {
                        int ColorIndex = Format.FormatCell.bBackground;
                        if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                            bgColor = Color.FromArgb((int)( a1CPallete[ColorIndex] + 0xFF000000));
                    }
                    else
                        MakeTransparent = true;

                    g.Clear(bgColor);
                    HandleRef hdcsrc = new HandleRef(g, g.GetHdc());
                    result = OLE32.OleDraw(pOle, 1, hdcsrc, ref rect);
                    g.ReleaseHdc(hdcsrc.Handle);
                    if (MakeTransparent)
                        m.MakeTransparent(bgColor);

                    //m.Save($"L:\\OleObject_{dwItemNumber}.png", ImageFormat.Png);
                }
                //pOle.Close(0);
                //Marshal.ReleaseComObject(pOle);
                Marshal.ReleaseComObject(RootStorage);
                Marshal.ReleaseComObject(LockBytes);
                Marshal.FreeHGlobal(hGlobal);

                OLE32.CoUninitialize();
                return m;
            }

            private object LoadRectangle(Picture picture, BinaryReader br)
            {
                return null;
            }

            private Image LoadPicture(Picture picture, BinaryReader br)
            {
                uint xz = br.ReadUInt32();
                int PictureSize = br.ReadInt32();
                byte[] pictureBuffer = br.ReadBytes(PictureSize);

                Bitmap Pic = Image.FromStream(new MemoryStream(pictureBuffer)) as Bitmap;
                bool MakeTransparent = false;

                if (!Format.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                    MakeTransparent = true;

                if (MakeTransparent)
                    Pic.MakeTransparent(Color.White);

                return Pic;
            }

            private object LoadLine(Picture picture, DataCell format, BinaryReader br)
            {
                return null;
            }
            private object LoadTextBox(Picture picture, DataCell format, BinaryReader br)
            {
                return null;
            }


        }

        public class MoxelRow
        {
            public Cellv6 FormatCell = Cellv6.Empty;
            public Dictionary<int, DataCell> Cells = new Dictionary<int, DataCell>();
            Moxel Parent = null;

            public MoxelRow()
            { }

            public MoxelRow(BinaryReader br, Moxel parent)
            {
                Parent = parent;
                FormatCell = br.Read<DataCell>(parent);
                Cells = br.ReadDictionary<DataCell>(parent);
            }

        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        public struct Picture
        {
            public ObjectType dwType;
            public int dwColumnStart;
            public int dwRowStart;
            public int dwOffsetLeft;
            public int dwOffsetTop;
            public int dwColumnEnd;
            public int dwRowEnd;
            public int dwOffsetRight;
            public int dwOffsetBottom;
            public int dwZOrder;
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        struct CellsUnion
        {
            public int dwLeft;
            public int dwTop;
            public int dwRight;
            public int dwBottom;
            public static CellsUnion Empty = new CellsUnion();
            public string HtmlSpan
            {
                get
                {
                    string Span = string.Empty;

                    if (dwTop != dwBottom)
                        Span += $" rowspan=\"{dwBottom - dwTop + 1}\"";

                    if (dwRight != dwLeft)
                        Span += $" colspan=\"{dwRight - dwLeft + 1}\"";
                    return Span;
                }
            }

            public int ColumnSpan
            {
                get
                {
                    return dwRight - dwLeft;
                }

            }

            public int RowSpan
            {
                get
                {
                    return dwBottom - dwTop;
                }

            }

            public bool ContainsCell(int row, int column)
            {
                return dwRight >= column && dwLeft <= column && dwTop < row && dwBottom >= row;
            }
        };
        
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        public struct Area
        {
            public int Unknown1; // always 1
            public int Unknown2; // garbage
            public int AreaType;
            public int ColumnBegin;
            public int RowBegin;
            public int ColumnEnd;
            public int RowEnd;
        }; 

        public class MoxelArea
        {
            public string Name;
            public Area Area;

            public MoxelArea()
            { }

            public MoxelArea(BinaryReader br)
            {
                Name = br.ReadCString();
                Area = br.Read<Area>();
            }
        }

        public class Section
        {
            int Begin;
            int End;
            int Level;
            string Name;

            public Section()
            { }

            public Section(BinaryReader br)
            {
                Begin = br.ReadInt32();
                End = br.ReadInt32();
                Level = br.ReadInt32();
                Name = br.ReadCString();
            }
        }


        public void Load(byte[] buf)
        {
            MemoryStream ms = new MemoryStream(buf);
            BinaryReader br = new BinaryReader(ms);
            stringTable = new Dictionary<int, string>();

            int nPos = 0xb;

            br.BaseStream.Seek(nPos, SeekOrigin.Begin);
            Version = br.ReadInt16();

            //Всего колонок
            nAllColumnCount = br.ReadInt32();
            //Всего строк
            nAllRowCount = br.ReadInt32();
            //Всего объектов
            nAllObjectsCount = br.ReadInt32();
            DefFormat = br.Read<DataCell>(this);
            FontList = br.ReadDictionary<LOGFONT>();

            int[] strnums = br.ReadIntArray();
            int stlCount = br.ReadCount();
            foreach (int num in strnums)
                stringTable.Add(num, br.ReadCString());

            TopColon = br.Read<DataCell>(this);
            BottomColon = br.Read<DataCell>(this);

            Columns = br.ReadDictionary<DataCell>(this);
            Rows = br.ReadDictionary<MoxelRow>(this);
            Objects = br.ReadList<EmbeddedObject>(this);
            Unions = br.ReadList<CellsUnion>();
            VerticalSections = br.ReadList<Section>();
            HorisontalSections = br.ReadList<Section>();
            HorisontalPageBreaks = br.ReadIntArray();
            VerticalPageBreaks = br.ReadIntArray();
            AreaNames = br.ReadList<MoxelArea>();

        }

        public static readonly uint[] a1CPallete =
            {
        //		0xRRGGBB,
		        0x000000,
                0xFFFFFF,
                0xFF0000,
                0x00FF00,
                0x0000FF,
                0xFFFF00,
                0xFF00FF,
                0x00FFFF,

                0x800000,
                0x008000,
                0x808000,
                0x000080,
                0x800080,
                0x008080,
                0x808080,
                0xC0C0C0,

                0x8080FF,
                0x802060,
                0xFFFFC0,
                0xA0E0E0,
                0x600080,
                0xFF8080,
                0x0080C0,
                0xC0C0FF,

                0x00CFFF,
                0x69FFFF,
                0xE0FFED,
                0xDD9CB3,
                0xB38FEE,
                0x2A6FF9,
                0x3FB8CD,
                0x488436,

                0x958C41,
                0x8E5E42,
                0xA0627A,
                0x624FAC,
                0x1D2FBE,
                0x286676,
                0x004500,
                0x453E01,

                0x6A2813,
                0x85396A,
                0x4A3285,
                0xC0DCC0,
                0xA6CAF0,
                0x800000,
                0x008000,
                0x000080,

                0x808000,
                0x800080,
                0x008080,
                0x808080,
                0xFFFBF0,
                0xA0A0A4,
                0x313900,
                0xD98534
            };

        uint GetPallete(int nIndex)
        {

            if (nIndex < 0 || nIndex >= a1CPallete.Length)
                return 0;

            uint nRes = a1CPallete[nIndex];
            return nRes; // RGB(nRes >> 16 & 0xFF, nRes >> 8 & 0xFF, nRes & 0xFF);
        }

        public enum BorderStyle : byte
        {
            None = 0,
            ThinDotted = 1,
            ThinSolid = 2,
            MediumSolid = 3,
            ThickSolid = 4,
            Double = 5,
            ThinDashedShort = 6,
            ThinDashedLong = 7,
            ThinGrayDotted = 8,
            MediumDashed = 9
        };


        public enum ObjectBorderStyle : byte
        {
            None = 0,
            Solid = 1,
            DashedExtraLong = 2,
            DashedShort = 3,
            DashDotSparse = 4,
            DashDotDot = 5
        };

        public enum ObjectBorderWidth : byte
        {
            Thin = 0,
            Medium = 1,
            Thick = 2
        };

        [Flags]
        public enum ObjectBorderPresence : byte
        {
            Left = 0x01,
            Top = 0x02,
            Right = 0x04,
            Bottom = 0x08,
            All = 0x0F,
        }

         
        public enum clFontWeight : byte
        {
            Normal = 0x04,
            Bold = 0x07
        };


    }
}
