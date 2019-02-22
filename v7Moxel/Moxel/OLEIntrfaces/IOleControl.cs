using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [StructLayout(LayoutKind.Sequential)]
    public struct ACCEL
    {
        byte fVirt;
        [MarshalAs(UnmanagedType.U2)]
        uint key;
        [MarshalAs(UnmanagedType.U2)]
        uint cmd;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CONTROLINFO
    {
        [MarshalAs(UnmanagedType.U4)]
        int cb;
        [MarshalAs(UnmanagedType.Struct)]
        ACCEL hAccel;
        [MarshalAs(UnmanagedType.U2)]
        uint cAccel;
        [MarshalAs(UnmanagedType.U4)]
        uint dwFlags;
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("B196B288-BAB4-101A-B69C-00AA00341D07")]
    public interface IOleControl
    {
        [PreserveSig]
        void GetControlInfo([In, Out] ref CONTROLINFO pCI);
        [PreserveSig]
        void OnMnemonic([In]System.Windows.Forms.Message pMsg);
        [PreserveSig]
        void OnAmbientPropertyChange([In]int dispID);
        [PreserveSig]
        void FreezeEvents([In, MarshalAs(UnmanagedType.Bool)] bool bFreeze);
    }
}
