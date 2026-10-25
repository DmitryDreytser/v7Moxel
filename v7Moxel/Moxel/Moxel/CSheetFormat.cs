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
    // ============================================================================
    //  CSheetFormat.cs — исправленная копия (Task 7).
    //  Изменения (API не менялся, раскладка 30 байт не менялась):
    //  1. БАГ: в цветовых геттерах проверка границ палитры была
    //     "bX >= 0 || bX < a1CPallete.Length" (в BgColor) и
    //     "bX > 0 || bX < a1CPallete.Length" (в остальных) — условие с "||"
    //     всегда истинно, поэтому индексы палитры >= 56 падали с
    //     IndexOutOfRangeException. Заменено на честную проверку "bX < длины".
    //     Для всех валидных индексов результат идентичен прежнему
    //     (индекс 0 у BorderColor/FontColor раньше давал Color.Black,
    //     palette[0] — тоже чёрный).
    //  2. GetHashCode: был маршалинг структуры через AllocHGlobal на КАЖДЫЙ
    //     вызов + XOR всех байт (порядконезависимая свёртка — массовые
    //     коллизии). Заменён на дешёвый хэш по полям без аллокаций.
    // ============================================================================

    [Serializable]
    [StructLayout(LayoutKind.Explicit, CharSet = CharSet.Ansi, Pack = 1, Size = 30)]
    unsafe public struct CSheetFormat
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
                              // union{
        [FieldOffset(0x06)]
        public short wStartPage; // Колонтитулы
        [FieldOffset(0x06)]
        public short wWidth; // Колонки
        [FieldOffset(0x06)]
        public short wRowPosition; // Строки
                                   // }
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
        // };
        // union {
        [FieldOffset(0x13)]
        public BorderStyle bBorderTop;
        [FieldOffset(0x13)]
        public ObjectBorderWidth bPictureBorderWidth;
        // };
        // union {
        [FieldOffset(0x14)]
        public BorderStyle bBorderRight;
        [FieldOffset(0x14)]
        public ObjectBorderPresence bPictureBorderPresence;
        // };
        // union {
        [MarshalAs(UnmanagedType.U1)]
        [FieldOffset(0x15)]
        public bool bPrintPicture;
        [FieldOffset(0x15)]
        public BorderStyle bBorderBottom;
        // };
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

        public static CSheetFormat Empty = new CSheetFormat
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
                if (dwFlags.HasFlag(MoxelCellFlags.BorderColor) &&
                    bBorderColor < a1CPallete.Length)
                    return Color.FromArgb((int)(a1CPallete[bBorderColor] + 0xFF000000));
                return Color.Black;
            }
        }

        public Color BgColor
        {
            get
            {
                if (dwFlags.HasFlag(MoxelCellFlags.Background) &&
                    bBackground < a1CPallete.Length)
                    return Color.FromArgb((int)(a1CPallete[bBackground] + 0xFF000000));
                return Color.Empty;
            }
        }

        public Color PatternColor
        {
            get
            {
                if (dwFlags.HasFlag(MoxelCellFlags.PatternColor) &&
                    bPatternColor < a1CPallete.Length)
                    return Color.FromArgb((int)(a1CPallete[bPatternColor] + 0xFF000000));
                return Color.Black;
            }
        }

        public Color FontColor
        {
            get
            {
                if (dwFlags.HasFlag(MoxelCellFlags.FontColor) &&
                    bFontColor < a1CPallete.Length)
                    return Color.FromArgb((int)(a1CPallete[bFontColor] + 0xFF000000));
                return Color.Black;
            }
        }

        public override int GetHashCode()
        {
            // дешёвый хэш по полям: без Marshal.AllocHGlobal на каждый вызов;
            // прежний XOR байт был порядконезависимым — перестановки полей
            // давали один и тот же хэш
            unchecked
            {
                var h = (int)dwFlags;
                h = (h * 397) ^ wHeight;
                h = (h * 397) ^ wWidth;
                h = (h * 397) ^ wStartPage;
                h = (h * 397) ^ wFontNumber;
                h = (h * 397) ^ wFontSize;
                h = (h * 397) ^ (byte)bFontBold;
                h = (h * 397) ^ (bFontItalic ? 1 : 0);
                h = (h * 397) ^ (bFontUnderline ? 1 : 0);
                h = (h * 397) ^ (int)bHorAlign;
                h = (h * 397) ^ (int)bVertAlign;
                h = (h * 397) ^ bPatternType;
                h = (h * 397) ^ (int)bBorderLeft;
                h = (h * 397) ^ (int)bBorderTop;
                h = (h * 397) ^ (int)bBorderRight;
                h = (h * 397) ^ (int)bBorderBottom;
                h = (h * 397) ^ bPatternColor;
                h = (h * 397) ^ bBorderColor;
                h = (h * 397) ^ bFontColor;
                h = (h * 397) ^ bBackground;
                h = (h * 397) ^ (int)bControlContent;
                h = (h * 397) ^ (int)bType;
                h = (h * 397) ^ (bAllowEdit ? 1 : 0);
                h = (h * 397) ^ bXZ1;
                return h;
            }
        }
    }
}
