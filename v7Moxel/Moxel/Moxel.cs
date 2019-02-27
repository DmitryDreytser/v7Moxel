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
using System.Collections;

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

        public Cellv6 DefFormat;
        public Dictionary<int, LOGFONT> FontList;
        public Dictionary<int, string> stringTable;
        public DataCell TopColon;
        public DataCell BottomColon;

        public Dictionary<int, DataCell> Columns;
        public Dictionary<int, MoxelRow> Rows;
        public List<EmbeddedObject> Objects;
        public List<CellsUnion> Unions;
        public List<Section> VerticalSections;
        public List<Section> HorisontalSections;
        public int[] HorisontalPageBreaks;
        public int[] VerticalPageBreaks;
        public List<MoxelArea> AreaNames;



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

            if(Rows.ContainsKey(RowNumber))
                result = Rows[RowNumber].Height;

            if (result == 0)
                if (DefFormat.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                    result = DefFormat.wHeight;

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
                dwFlags = 0,
                wShow = 0,
                wColumnPosition = 0,
                wHeight = 0,
                wStartPage = 0,
                wWidth = 0,
                wRowPosition = 0,
                wFontNumber = 0,
                wFontSize = 0,
                bFontBold = 0,
                bFontItalic = false,
                bFontUnderline = false,
                bHorAlign = 0,
                bVertAlign = 0,
                bPatternType = 0,
                bBorderLeft = 0,
                bPictureBorderStyle = 0,
                bBorderTop = 0,
                bPictureBorderWidth = 0,
                bBorderRight = 0,
                bPictureBorderPresence = 0,
                bBorderBottom = 0,
                bPrintPicture = false,
                bPatternColor = 0,
                bBorderColor = 0,
                bFontColor = 0,
                bBackground = 0,
                bControlContent = 0,
                bType = 0,
                bAllowEdit = false,
                bXZ1 = 0
            };

            public Color BorderColor
            {
                get
                {
                    if (dwFlags.HasFlag(MoxelCellFlags.BorderColor))
                        if (bBorderColor > 0 || bBorderColor < a1CPallete.Length)
                            return Color.FromArgb((int)(a1CPallete[bBorderColor] + 0xFF000000));
                    return Color.Black;
                }
            }

            public Color BgColor
            {
                get
                {
                    if (dwFlags.HasFlag(MoxelCellFlags.Background))
                        if (bBackground >= 0 || bBackground < a1CPallete.Length)
                            return Color.FromArgb((int)(a1CPallete[bBackground] + 0xFF000000));
                    return Color.Empty;
                }
            }

            public Color PatternColor
            {
                get
                {
                    if (dwFlags.HasFlag(MoxelCellFlags.PatternColor))
                        if (bPatternColor > 0 || bPatternColor < a1CPallete.Length)
                            return Color.FromArgb((int)(a1CPallete[bPatternColor] + 0xFF000000));
                    return Color.Black;
                }
            }
            public Color FontColor
            {
                get
                {
                    if (dwFlags.HasFlag(MoxelCellFlags.FontColor))
                        if (bFontColor > 0 || bFontColor < a1CPallete.Length)
                            return Color.FromArgb((int)(a1CPallete[bFontColor] + 0xFF000000));
                    return Color.Black;
                }
            }
        }

        public class DataCell
        {
            public Cellv6 FormatCell;
            public string Text;
            public string Value;
            public byte[] Data;
            public short TextOrientation = 0;
            public Moxel Parent;

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

            public static implicit operator DataCell(Cellv6 c)
            {
                return new DataCell { FormatCell = c, TextOrientation = 0, Data = new byte[0], Parent = null};
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

        public class EmbeddedObject : DataCell
        {

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
            public Picture Picture;
            public object pObject;
            public object OleObject;
            public string ProgID;
            public Guid ClsId;
            public byte[] OleObjectStorage;

            public EmbeddedObject()
            { }
            public EmbeddedObject(BinaryReader br, Moxel parent) : base(br, parent)
            {
                Picture = br.Read<Picture>();
                switch (Picture.dwType)
                {
                    case ObjectType.Picture:
                        pObject = LoadPicture(br);
                        break;
                    case ObjectType.Ole:
                        pObject = LoadOleObject(br);
                        break;
                    default:
                        break; 
                }
            }

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

                    if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                    {
                        int ColorIndex = FormatCell.bBackground;
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
                }
                Marshal.ReleaseComObject(RootStorage);
                Marshal.ReleaseComObject(LockBytes);
                Marshal.FreeHGlobal(hGlobal);

                OLE32.CoUninitialize();
                return m;
            }
             
            private Image LoadPicture( BinaryReader br)
            {
                uint xz = br.ReadUInt32();
                int PictureSize = br.ReadInt32();
                byte[] pictureBuffer = br.ReadBytes(PictureSize);

                using (var memoryStream = new MemoryStream(pictureBuffer))
                {
                    Bitmap Pic = Image.FromStream(memoryStream) as Bitmap;
                    bool MakeTransparent = false;

                    if (!FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                        MakeTransparent = true;

                    if (MakeTransparent)
                        Pic.MakeTransparent(Color.White);

                    return Pic;
                }
            }
        }

        public class MoxelRow : IDictionary<int, DataCell>
        {
            Moxel Parent = null;
            public Cellv6 FormatCell;
            Dictionary<int, DataCell> values;

            public int Height
            {
                get
                {
                    if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                        if (FormatCell.wHeight == 0)
                            return 45;
                        else
                            return FormatCell.wHeight;
                    else
                        return 0;
                }
                set
                {
                    FormatCell.wHeight = (short)value;
                    if (value != 0)
                        FormatCell.dwFlags |= MoxelCellFlags.RowHeight;
                    else
                        FormatCell.dwFlags ^=  MoxelCellFlags.RowHeight;
                }
            }

            public ICollection<int> Keys
            {
                get
                {
                    return values.Keys;
                }
            }

            public ICollection<DataCell> Values
            {
                get
                {
                    return values.Values;
                }
            }

            public int Count
            {
                get
                {
                    return Parent.nAllColumnCount; 
                }
            }

            public bool IsReadOnly
            {
                get
                {
                    return false;
                }
            }

            public DataCell this[int key]
            {
                get
                {
                    if (values.ContainsKey(key))
                        return values[key];
                    else
                        return FormatCell;
                }

                set
                {
                   values[key] = value;
                }
            }

            public MoxelRow()
            { }

            public MoxelRow(BinaryReader br, Moxel parent)
            {
                Parent = parent;
                FormatCell = br.Read<DataCell>(parent);
                values = br.ReadDictionary<DataCell>(parent);
            }

            public bool ContainsKey(int key)
            {
                return values.ContainsKey(key);
            }

            public void Add(int key, DataCell value)
            {
                values.Add(key, value);
            }

            public bool Remove(int key)
            {
                return values.Remove(key);
            }

            public bool TryGetValue(int key, out DataCell value)
            {
                return values.TryGetValue(key, out value);
            }

            public IEnumerator GetEnumerator()
            {
                return Values.GetEnumerator();
            }

            public void Add(KeyValuePair<int, DataCell> item)
            {
                values.Add(item.Key, item.Value);
            }

            public void Clear()
            {
                values.Clear();
            }

            public bool Contains(KeyValuePair<int, DataCell> item)
            {
                return values.Contains(item);
            }

            public void CopyTo(KeyValuePair<int, DataCell>[] array, int arrayIndex)
            {
                ((ICollection)values).CopyTo(array, arrayIndex);
            }

            public bool Remove(KeyValuePair<int, DataCell> item)
            {
                return values.Remove(item.Key);
            }

            IEnumerator<KeyValuePair<int, DataCell>> IEnumerable<KeyValuePair<int, DataCell>>.GetEnumerator()
            {
                return values.GetEnumerator();
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
        public struct CellsUnion
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
