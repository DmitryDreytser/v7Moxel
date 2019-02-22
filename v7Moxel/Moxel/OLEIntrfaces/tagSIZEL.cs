using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{

    [Serializable]
    [StructLayout(LayoutKind.Sequential)]
    public struct tagSIZEL
    {
        public int cx;
        public int cy;
        //public tagSIZEL() { }
        public tagSIZEL(int cx, int cy) { this.cx = cx; this.cy = cy; }
        public tagSIZEL(tagSIZEL o) { this.cx = o.cx; this.cy = o.cy; }
    }
}
