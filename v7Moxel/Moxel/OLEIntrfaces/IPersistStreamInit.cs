using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;

namespace Ole
{
    [ComImport]
    [Guid("7FD52380-4E07-101B-AE2D-08002B2EC713")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPersistStreamInit : IPersist
    {
        [PreserveSig]
        int GetSizeMax([Out, MarshalAs(UnmanagedType.I8)] out long pCbSize);

        [PreserveSig]
        int InitNew();

        [PreserveSig]
        [return: MarshalAs(UnmanagedType.Bool)]
        bool IsDirty();

        [PreserveSig]
        int Load([In, MarshalAs(UnmanagedType.Interface)] IStream pStm);

        [PreserveSig]
        int Save([In, MarshalAs(UnmanagedType.Interface)] IStream pStm, [In, MarshalAs(UnmanagedType.Bool)] bool fClearDirty);

    }
}
