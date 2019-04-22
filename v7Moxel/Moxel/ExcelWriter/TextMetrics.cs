using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Moxel
{
    static class TextMetrics
    {
        static class NativeMethods
        {

            [DllImport("gdi32.dll", CharSet = CharSet.Auto)]
            public static extern bool GetTextMetrics(IntPtr hdc, out TEXTMETRICW lptm);

            [DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern bool DeleteObject(IntPtr hObject);

            [DllImport("gdi32.dll", CharSet = CharSet.Auto)]
            public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiObj);

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            public struct TEXTMETRICW
            {
                public int tmHeight;
                public int tmAscent;
                public int tmDescent;
                public int tmInternalLeading;
                public int tmExternalLeading;
                public int tmAveCharWidth;
                public int tmMaxCharWidth;
                public int tmWeight;
                public int tmOverhang;
                public int tmDigitizedAspectX;
                public int tmDigitizedAspectY;
                public ushort tmFirstChar;
                public ushort tmLastChar;
                public ushort tmDefaultChar;
                public ushort tmBreakChar;
                public byte tmItalic;
                public byte tmUnderlined;
                public byte tmStruckOut;
                public byte tmPitchAndFamily;
                public byte tmCharSet;
            }
        }

        public static int GetFontAverageCharWidth(IDeviceContext dc, Font font)
        {
            int result;
            IntPtr hDC;
            IntPtr hFont;
            IntPtr hFontDefault;

            hDC = IntPtr.Zero;
            hFont = IntPtr.Zero;
            hFontDefault = IntPtr.Zero;

            try
            {
                NativeMethods.TEXTMETRICW textMetric;

                hDC = dc.GetHdc();

                hFont = font.ToHfont();
                hFontDefault = NativeMethods.SelectObject(hDC, hFont);

                NativeMethods.GetTextMetrics(hDC, out textMetric);

                result = textMetric.tmAveCharWidth;
            }

            finally
            {
                if (hFontDefault != IntPtr.Zero)
                {
                    NativeMethods.SelectObject(hDC, hFontDefault);
                }

                if (hFont != IntPtr.Zero)
                {
                    NativeMethods.DeleteObject(hFont);
                }

                dc.ReleaseHdc();
            }

            return result;
        }
    }
}
