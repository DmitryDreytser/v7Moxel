using System;
using System.Runtime.InteropServices;

namespace Moxel
{
    [Serializable]
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
    public struct CellsUnion
    {
        public int dwLeft;
        public int dwTop;
        public int dwRight;
        public int dwBottom;
        public static CellsUnion Empty = new CellsUnion();

        public bool IsEmpty()
        {
            return dwLeft == 0 && dwRight == 0 && dwTop == 0 && dwBottom == 0;
        }

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

}
