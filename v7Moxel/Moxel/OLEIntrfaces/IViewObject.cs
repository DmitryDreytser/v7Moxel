using System;
using System.Drawing;
using System.Runtime.InteropServices;
using IAdviseSink = System.Runtime.InteropServices.ComTypes.IAdviseSink;


namespace Ole
{
    [Guid("0000010d-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport()]
    public interface IViewObject
    {
        HRESULT Draw([MarshalAs(UnmanagedType.U4)] int dwDrawAspect, int lindex, IntPtr pvAspect, DVTARGETDEVICE ptd, IntPtr hdcTargetDev, HandleRef hdcDraw, ref Rectangle lprcBounds, ref Rectangle lprcWBounds, IntPtr pfnContinue, int dwContinue);
        //void Draw([MarshalAs(UnmanagedType.U4)] int dwDrawAspect, int lindex, int pvAspect, int ptd, int hdcTargetDev, IntPtr hdcDraw, ref COMRECT lprcBounds, int lprcWBounds, int pfnContinue, int dwContinue);
        HRESULT GetColorSet([MarshalAs(UnmanagedType.U4)] int dwDrawAspect, int lindex, IntPtr pvAspect, DVTARGETDEVICE ptd, IntPtr hicTargetDev, out tagLOGPALETTE ppColorSet);
        HRESULT Freeze([MarshalAs(UnmanagedType.U4)] int dwDrawAspect, int lindex, IntPtr pvAspect, out IntPtr pdwFreeze);
        HRESULT Unfreeze([MarshalAs(UnmanagedType.U4)] int dwFreeze);
        HRESULT SetAdvise([MarshalAs(UnmanagedType.U4)] int aspects, [MarshalAs(UnmanagedType.U4)] int advf, [MarshalAs(UnmanagedType.Interface)] IAdviseSink pAdvSink);
        void GetAdvise([MarshalAs(UnmanagedType.LPArray)] out int[] paspects, [MarshalAs(UnmanagedType.LPArray)] out int[] advf, [MarshalAs(UnmanagedType.LPArray)] out IAdviseSink[] pAdvSink);
    }

    public struct RECTL
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }
    

    [StructLayout(LayoutKind.Sequential)]

    public class COMRECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
        public COMRECT() { }
        public COMRECT(Rectangle r)
        {
            this.left = r.X;
            this.top = r.Y;
            this.right = r.Right;
            this.bottom = r.Bottom;
        }

        public COMRECT(int left, int top, int right, int bottom)
        {
            this.left = left;
            this.top = top;
            this.right = right;
            this.bottom = bottom;
        }

        public static COMRECT FromXYWH(int x, int y, int width, int height)

        {

            return new COMRECT(x, y, x + width, y + height);

        }

        public override string ToString()

        {

            return string.Concat(new object[] { "Left = ", this.left, " Top ", this.top, " Right = ", this.right, " Bottom = ", this.bottom });

        }

    }

    [StructLayout(LayoutKind.Sequential)]

    public sealed class tagLOGPALETTE

    {

        [MarshalAs(UnmanagedType.U2)]

        public short palVersion;

        [MarshalAs(UnmanagedType.U2)]

        public short palNumEntries;

        public tagLOGPALETTE() { }

    }

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public class DVTARGETDEVICE
    {
        [MarshalAs(UnmanagedType.U4)]
        public int tdSize;
        [MarshalAs(UnmanagedType.U2)]
        public short tdDriverNameOffset;
        [MarshalAs(UnmanagedType.U2)]
        public short tdDeviceNameOffset;
        [MarshalAs(UnmanagedType.U2)]
        public short tdPortNameOffset;
        [MarshalAs(UnmanagedType.U2)]
        public short tdExtDevmodeOffset;
        //[MarshalAs(UnmanagedType.LPArray, SizeConst = 1)]
        byte tdData;
    }
}