using System;
using System.Drawing;
using System.Linq;

namespace Ole
{
    public class RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;

        public int Width { get { return right - left; } }
        public int Height { get { return bottom - top; } }

        public static implicit operator RECT(Rectangle rect)
        {
            return new RECT { left = rect.Left, right = rect.Right, top = rect.Top, bottom = rect.Bottom };
        }
    }
}
