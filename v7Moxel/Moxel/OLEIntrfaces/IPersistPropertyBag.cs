using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [ComImport]
    [Guid("37D84F60-42CB-11CE-8135-00AA004BB851")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPersistPropertyBag
    {
        [PreserveSig]
        int InitNew();

        void Load(
            [In, MarshalAs(UnmanagedType.Interface)] IPropertyBag pPropBag,
            [In, MarshalAs(UnmanagedType.Interface)] IErrorLog pErrorLog);

        void Save([In, MarshalAs(UnmanagedType.Interface)] IPropertyBag pPropBag,
            [In, MarshalAs(UnmanagedType.Bool)] bool fClearDirty,
            [In, MarshalAs(UnmanagedType.Bool)] bool fSaveAllProperties);

    }
}
