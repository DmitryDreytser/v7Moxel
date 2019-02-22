using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [ComConversionLoss]
    // IOleClientSite
    [ComImport(), Guid("00000118-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleClientSite
    {
        void SaveObject();

        [return: MarshalAs(UnmanagedType.Interface)]
        object GetMoniker(
            [In, MarshalAs(UnmanagedType.U4)] int dwAssign,
            [In, MarshalAs(UnmanagedType.U4)] int dwWhichMoniker);

        [PreserveSig]
        int GetContainer([Out] out IntPtr ppContainer);

        void ShowObject();

        void OnShowWindow([In, MarshalAs(UnmanagedType.Bool)] bool fShow);

        void RequestNewObjectLayout();
    }
}
