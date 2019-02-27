using System.Runtime.InteropServices;
using static Moxel.Moxel;

namespace Moxel
{
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

}
