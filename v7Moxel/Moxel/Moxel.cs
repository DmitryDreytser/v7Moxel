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

        Dictionary<int, Cellv6> Columns;
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
                if(Columns[ColNumber].dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                    result = Columns[ColNumber].wWidth;

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
            if (DefFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                result = DefFormat.wHeight;

            if (Rows.ContainsKey(RowNumber))
                if (Rows[RowNumber].FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
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

        public bool SaveToHtml(string filename)
        {
            Cellv6 FormatCell = DefFormat;
            string DefFontNAme = string.Empty;
            float DefFontSize = 8.0f;

            if (FontList.Count == 1)
                DefFontNAme = $" style=\"font-family:{FontList.Last().Value.lfFaceName}\"";

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
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(Text))
                    {
                        string FontFamily = DefFontNAme;
                        float FontSize = DefFontSize;

                        FillFormat(FormatCell, ref CellStyle, ref Text, ref FontFamily, ref FontSize);

                        Text = Text.Replace("\r\n", "<br>");
                        
                        Size Constr = new Size { Width = (int)Math.Round((GetWidth(columnnumber, columnnumber + Union.ColumnSpan) + GetColumnWidth(columnnumber)) * 0.875), Height = 0 };
                        Size textsize = System.Windows.Forms.TextRenderer.MeasureText(Text.TrimStart(' '), new Font(FontFamily, FontSize, FormatCell.bFontBold == clFontWeight.Bold ? FontStyle.Bold : FontStyle.Regular), Constr, TextFormatFlags.WordBreak | TextFormatFlags.ExternalLeading);

                        textsize.Height /= Union.RowSpan + 1;

                        if (!Row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
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

                if (Row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
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
                FormatCell = obj.Format.FormatCell;

                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                {
                    int ColorIndex = FormatCell.bBackground;
                    if (ColorIndex >= 0 || ColorIndex < a1CPallete.Length)
                    {
                        Color bgColor = Color.FromArgb((int)(a1CPallete[ColorIndex] + 0xFF000000));
                        PictureStyle.Set("background-color", $"rgb({bgColor.R},{bgColor.G},{bgColor.B})");
                    }
                }


                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
                {
                    if(FormatCell.bPictureBorderPresence != ObjectBorderPresence.All)
                    {
                        if(FormatCell.bPictureBorderPresence.HasFlag(ObjectBorderPresence.Left))
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

                    PictureStyle.Set("border-width", $"{(byte)FormatCell.bPictureBorderWidth + 1}px");
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

                int StartPx = (int)Math.Round(GetWidth(0, obj.Picture.dwColumnStart) * 0.875 + obj.Picture.dwOffsetLeft / 2.8,0);
                PictureStyle.Add("left", $"{StartPx}px");
                int Width = (int)Math.Round(GetWidth(0, obj.Picture.dwColumnEnd) * 0.875 + obj.Picture.dwOffsetRight / 2.8 - StartPx, 0);

                PictureStyle.Add("width", $"{Width}px");

                int Top = (int)Math.Round(GetHeight(0, obj.Picture.dwRowStart) / 3 + obj.Picture.dwOffsetTop / 2.8, 0);
                PictureStyle.Add("Top", $"{Top}px");

                int Height = (int)Math.Round(GetHeight(0, obj.Picture.dwRowEnd) / 3 + obj.Picture.dwOffsetBottom / 2.8 - Top, 0);
                PictureStyle.Add("height", $"{Height}px");


                PictureStyle.Add("position", "absolute");
                result.Append($"\t\t<div id=\"D{obj.dwItemNumber}\"{PictureStyle}>\r\n");


                string DataUri = string.Empty;
                switch (obj.Picture.dwType)
                {
                    case ObjectType.Ole:
                    case ObjectType.Picture:
                            using (MemoryStream ms = new MemoryStream())
                            {
                                ((Bitmap)obj.pObject)?.Save(ms, ImageFormat.Png);
                                DataUri = Convert.ToBase64String(ms.ToArray());
                                result.Append($"\t\t\t<img src=\"data:image/png;base64,{DataUri}\" width=\"{Width}\" height=\"{Height}\">\r\n");
                            }
                        break;
                    case ObjectType.Text:
                        CSSstyle TextStyle = new CSSstyle();
                        FillFormat(obj.Format.FormatCell, ref TextStyle, ref obj.Format.Text);
                        result.AppendLine($"<p{TextStyle}>{obj.Format.Text}</p>");
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

            public DataCell()
            { }
            public DataCell(BinaryReader br)
            {
                FormatCell = br.Read<Cellv6>();
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

            public DataCell Format;
            public Picture Picture;
            public object pObject;
            public object OleObject;
            public string ProgID;
            public Guid ClsId;
            public int dwItemNumber = 0;

            public EmbeddedObject()
            { }
            public EmbeddedObject(BinaryReader br)
            {
                Format = br.Read<DataCell>();
                Picture = br.Read<Picture>();
                switch (Picture.dwType)
                {
                    case ObjectType.Picture:
                        pObject = LoadPicture(Picture, br );
                        break;
                    case ObjectType.Line:
                        pObject = LoadLine(Picture, Format, br);
                        break;
                    case ObjectType.Rectangle:
                        pObject = LoadRectangle(Picture, br);
                        break;
                    case ObjectType.Text:
                        pObject = LoadTextBox(Picture, Format, br);
                        break;
                    case ObjectType.Ole:
                        pObject = LoadOleObject(br);
                        break;
                    default:
                        throw new Exception("Неизвестный тип внедренного объекта");
                }
            }
            enum TernaryRasterOperations : uint
            {
                /// <summary>dest = source</summary>
                SRCCOPY = 0x00CC0020,
                /// <summary>dest = source OR dest</summary>
                SRCPAINT = 0x00EE0086,
                /// <summary>dest = source AND dest</summary>
                SRCAND = 0x008800C6,
                /// <summary>dest = source XOR dest</summary>
                SRCINVERT = 0x00660046,
                /// <summary>dest = source AND (NOT dest)</summary>
                SRCERASE = 0x00440328,
                /// <summary>dest = (NOT source)</summary>
                NOTSRCCOPY = 0x00330008,
                /// <summary>dest = (NOT src) AND (NOT dest)</summary>
                NOTSRCERASE = 0x001100A6,
                /// <summary>dest = (source AND pattern)</summary>
                MERGECOPY = 0x00C000CA,
                /// <summary>dest = (NOT source) OR dest</summary>
                MERGEPAINT = 0x00BB0226,
                /// <summary>dest = pattern</summary>
                PATCOPY = 0x00F00021,
                /// <summary>dest = DPSnoo</summary>
                PATPAINT = 0x00FB0A09,
                /// <summary>dest = pattern XOR dest</summary>
                PATINVERT = 0x005A0049,
                /// <summary>dest = (NOT dest)</summary>
                DSTINVERT = 0x00550009,
                /// <summary>dest = BLACK</summary>
                BLACKNESS = 0x00000042,
                /// <summary>dest = WHITE</summary>
                WHITENESS = 0x00FF0062,
                /// <summary>
                /// Capture window as seen on screen.  This includes layered windows 
                /// such as WPF windows with AllowsTransparency="true"
                /// </summary>
                CAPTUREBLT = 0x40000000
            }

            [DllImport("user32.dll", SetLastError = false)]
            static extern IntPtr GetDesktopWindow();

            [DllImport("user32.dll", SetLastError = false)]
            static extern IntPtr GetWindowDC(IntPtr hWnd);

            [DllImport("gdi32.dll",  SetLastError = true)]
            static extern int GetDeviceCaps(IntPtr hDC, int Caps);

            [DllImport("gdi32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            static extern bool BitBlt([In] IntPtr hdc, int nXDest, int nYDest, int nWidth, int nHeight, [In] IntPtr hdcSrc, int nXSrc, int nYSrc, TernaryRasterOperations dwRop);

            [DllImport("gdi32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            static extern bool StretchBlt(IntPtr hdc,  int xDest,  int yDest,  int wDest,  int hDest, [In] IntPtr hdcSrc,  int xSrc,  int ySrc,  int wSrc,  int hSrc,  TernaryRasterOperations dwRop);


            public static Bitmap CopyGraphicsContent(Graphics source, Rectangle rect)
            {
                Bitmap bmp = new Bitmap(rect.Width, rect.Height);
                using (Graphics dest = Graphics.FromImage(bmp))
                {
                    IntPtr hdcSource = source.GetHdc();
                    IntPtr hdcDest = dest.GetHdc();
                    BitBlt(hdcDest, 0, 0, rect.Width, rect.Height, hdcSource, rect.X, rect.Y, TernaryRasterOperations.SRCCOPY);
                    source.ReleaseHdc(hdcSource);
                    dest.ReleaseHdc(hdcDest);
                }
                return bmp;
            }

            [DllImport("ole32.dll")]
            static extern int ProgIDFromCLSID([In] ref Guid clsid,
             [MarshalAs(UnmanagedType.LPWStr)] out string lplpszProgID);


            [DllImport("ole32.dll")]
            static extern HRESULT OleSetContainedObject(
                  [In] IntPtr pUnknown,
                  [MarshalAs(UnmanagedType.Bool)]
                  [In] bool  fContained
                );



            [DllImport("ole32.dll")]
            static extern HRESULT OleNoteObjectVisible(
                  [In] IntPtr pUnknown,
                  [MarshalAs(UnmanagedType.Bool)]
                  [In] bool  fVisible);

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
                dwItemNumber = br.ReadInt32();
                int dwAspect = br.ReadInt32();
                short wUseMoniker = br.ReadInt16();

                dwAspect = br.ReadInt32();

                int dwObjectSize = br.ReadInt32();
                byte[] ObjectStorage = br.ReadBytes(dwObjectSize);

                //File.WriteAllBytes($"L:\\OleObject_{dwItemNumber}.bin", ObjectStorage);

                OLE32.CoInitializeEx(IntPtr.Zero, OLE32.CoInit.ApartmentThreaded); //COINIT_APARTMENTTHREADED
                OLE32.ILockBytes LockBytes;
                OLE32.IStorage RootStorage;

                IntPtr hGlobal = Marshal.AllocHGlobal(ObjectStorage.Length);
                Marshal.Copy(ObjectStorage, 0, hGlobal, ObjectStorage.Length);
                OLE32.CreateILockBytesOnHGlobal(hGlobal, false, out LockBytes);
                OLE32.IOleObject pOle = null;

                HRESULT result = OLE32.StgOpenStorageOnILockBytes(LockBytes, null, OLE32.STGM.STGM_READWRITE | OLE32.STGM.STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out RootStorage);
                
                System.Runtime.InteropServices.ComTypes.STATSTG MetaDataInfo;
                RootStorage.Stat(out MetaDataInfo, 0);
                ClsId = MetaDataInfo.clsid;
                ProgIDFromCLSID(ref ClsId, out ProgID);

                Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

                IOleClientSite ole_cs = null;
                result = OLE32.OleLoad(RootStorage, ref IID_IUnknown, ole_cs, out pOle);
                if (result != HRESULT.S_OK)
                {
                    int res = Marshal.GetLastWin32Error();
                    return null;
                }

                IntPtr pUnknwn = Marshal.GetIUnknownForObject(pOle);
                result = OleSetContainedObject(pUnknwn, true);
                result = OleNoteObjectVisible(pUnknwn, true);
                Marshal.Release(pUnknwn);
                result = OLE32.OleRun(pOle);

                tagSIZEL sizel = new tagSIZEL();
                pOle.GetExtent(1, ref sizel);
                float LogUnitsPerDevPixel_X = 1, LogUnitsPerDevPixel_Y = 1;

                //using (Graphics g = Graphics.FromImage(new Bitmap(1, 1)))
                //{
                //    HandleRef hdcsrc = new HandleRef(g, g.GetHdc());
                //    LogUnitsPerDevPixel_X = ((float)UNITS_PER_INCH) / ((float)GetDeviceCaps(hdcsrc.Handle, LOGPIXELSX));
                //    LogUnitsPerDevPixel_Y = ((float)UNITS_PER_INCH) / ((float)GetDeviceCaps(hdcsrc.Handle, LOGPIXELSY));
                //    g.ReleaseHdc(hdcsrc.Handle);
                //}

                Rectangle rect = new Rectangle(0, 0, (int)(sizel.cx / LogUnitsPerDevPixel_X), (int)(sizel.cy / LogUnitsPerDevPixel_Y));
                Bitmap m = new Bitmap(rect.Width, rect.Height);

                using (Graphics g = Graphics.FromImage(m))
                {
                    Color bgColor = Color.White;

                    
                    bool MakeTransparent = false;
                    if (ProgID == "BMP1C.Bmp1cCtrl.1")
                        if ((pOle as _DBmp_1c).GrMode == 1)
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
                br.ReadUInt32();
                int PictureSize = br.ReadInt32();
                byte[] pictureBuffer = br.ReadBytes(PictureSize);

                switch (PictureSize)
                {
                    case 2:
                        break;
                    case 4:
                        break;
                    default:
                        break;
                }
                return null;
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
            public Cellv6 FormatCell;
            public Dictionary<int, DataCell> Cells;

            public MoxelRow()
            { }

            public MoxelRow(BinaryReader br)
            {
                FormatCell = br.Read<Cellv6>();
                Cells = br.ReadDictionary<DataCell>();
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
            DefFormat = br.Read<Cellv6>();
            FontList = br.ReadDictionary<LOGFONT>();

            int[] strnums = br.ReadIntArray();
            int stlCount = br.ReadCount();
            foreach (int num in strnums)
                stringTable.Add(num, br.ReadCString());

            TopColon = br.Read<DataCell>();
            BottomColon = br.Read<DataCell>();

            Columns = br.ReadDictionary<Cellv6>();
            Rows = br.ReadDictionary<MoxelRow>();
            Objects = br.ReadList<EmbeddedObject>();
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
