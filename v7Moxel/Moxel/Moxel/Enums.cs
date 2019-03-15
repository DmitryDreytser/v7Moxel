using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Moxel
{
    partial class Moxel
    {
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

        public enum TextVertAlign : byte
        {
            Top = 0,
            Bottom = 8,
            Middle = 0x18
        };


        public enum TextControl : byte
        {
            Auto = 0,
            Cut = 1,
            Fill = 2,
            Wrap = 3,
            Red = 4,
            FillAndRed = 5
        };

        public enum ContentType : byte
        {
            Text = 0,
            Expression = 1,
            Pattern = 2,
            FixedPattern = 3
        };

        public enum ObjectType
        {
            None = 0,   //1-линия
            Line = 1,   //1-линия
            Rectangle,  //2-квадрат
            Text,       //3-блок текста (но без текста)
            Ole,        //4-ОЛЕ обьект (в т.ч. диаграмма 1С)
            Picture     //5-картинка
        };

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
            Empty = 0,
            Left = 0x01,
            Top = 0x02,
            Right = 0x04,
            Bottom = 0x08,
            All = 0x0F,
        }

        public enum clFontWeight : byte
        {
            Empty = 0,
            Normal = 0x04,
            Bold = 0x07
        };

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
    }
}
