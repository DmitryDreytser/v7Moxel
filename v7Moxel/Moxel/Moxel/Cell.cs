using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static Moxel.Moxel;

namespace Moxel
{
    [StructLayout(LayoutKind.Explicit, CharSet = CharSet.Ansi, Pack = 1)]
    public struct Cellv6
    {
        [FieldOffset(0x00)]
        public MoxelCellFlags dwFlags; // MoxcelCellFlags
                                       // union{
        [FieldOffset(0x04)]
        public short wShow; // 1 - да, 0xFFFF - нет. Используется в колонтитулах
        [FieldOffset(0x04)]
        public short wColumnPosition; // Используется в колонках
        [FieldOffset(0x04)]
        public short wHeight; // Используется в строках
                              //}
                              // union{
        [FieldOffset(0x06)]
        public short wStartPage; // Колонтитулы
        [FieldOffset(0x06)]
        public short wWidth; // Колонки
        [FieldOffset(0x06)]
        public short wRowPosition; // Строки
                                   //}
        [FieldOffset(0x08)]
        public short wFontNumber;
        [FieldOffset(0x0A)]
        public short wFontSize;
        [FieldOffset(0x0C)]
        public clFontWeight bFontBold;
        [MarshalAs(UnmanagedType.U1)]
        [FieldOffset(0x0D)]
        public bool bFontItalic;
        [MarshalAs(UnmanagedType.U1)]
        [FieldOffset(0x0E)]
        public bool bFontUnderline;
        [FieldOffset(0x0F)]
        public TextHorzAlign bHorAlign;
        [FieldOffset(0x10)]
        public TextVertAlign bVertAlign;
        [FieldOffset(0x11)]
        public byte bPatternType;
        // union {
        [FieldOffset(0x12)]
        public BorderStyle bBorderLeft;
        [FieldOffset(0x12)]
        public ObjectBorderStyle bPictureBorderStyle;
        //};
        // union {
        [FieldOffset(0x13)]
        public BorderStyle bBorderTop;
        [FieldOffset(0x13)]
        public ObjectBorderWidth bPictureBorderWidth;
        //};
        // union {
        [FieldOffset(0x14)]
        public BorderStyle bBorderRight;
        [FieldOffset(0x14)]
        public ObjectBorderPresence bPictureBorderPresence;
        //};
        // union {
        [MarshalAs(UnmanagedType.U1)]
        [FieldOffset(0x15)]
        public bool bPrintPicture;
        [FieldOffset(0x15)]
        public BorderStyle bBorderBottom;
        //};
        [FieldOffset(0x16)]
        public byte bPatternColor;
        [FieldOffset(0x17)]
        public byte bBorderColor;
        [FieldOffset(0x18)]
        public byte bFontColor;
        [FieldOffset(0x19)]
        public byte bBackground;
        [FieldOffset(0x1A)]
        public TextControl bControlContent; // MoxcelControlContent
        [FieldOffset(0x1B)]
        public ContentType bType; // MoxcelContentType
        [MarshalAs(UnmanagedType.U1)]
        [FieldOffset(0x1C)]
        public bool bAllowEdit;
        [FieldOffset(0x1D)]
        public byte bXZ1;

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
}
