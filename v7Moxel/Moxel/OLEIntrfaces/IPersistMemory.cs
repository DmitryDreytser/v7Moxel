using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;

namespace Ole
{
    [ComImport]
    [Guid("BD1AE5E0-A6AE-11CE-BD37-504200C10000")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPersistMemory : IPersist
    {
        [PreserveSig]
        int GetSizeMax([Out, MarshalAs(UnmanagedType.I8)] out long pCbSize);

        [PreserveSig]
        int InitNew();

        [PreserveSig]
        int IsDirty();

        [PreserveSig]
        int Load([In, MarshalAs(UnmanagedType.Interface)] IStream pStm);

        [PreserveSig]
        int Save([In, MarshalAs(UnmanagedType.Interface)] IStream pStm, [In, MarshalAs(UnmanagedType.Bool)] bool fClearDirty);
    }
}
