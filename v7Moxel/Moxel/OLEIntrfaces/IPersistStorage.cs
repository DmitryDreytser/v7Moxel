using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [Guid("0000010A-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IPersistStorage //we don't need the base interface IPersist
    {
        [PreserveSig]
        void GetClassID(
            out Guid pClassID);
        [PreserveSig]
        [return: MarshalAs(UnmanagedType.Bool)]
        bool IsDirty();
        void InitNew(
            [MarshalAs(UnmanagedType.Interface)] [In] IStorage pstg);
        [PreserveSig]
        uint Load(
            [MarshalAs(UnmanagedType.Interface)] [In] IStorage pstg);
        [PreserveSig]
        uint Save(
            [MarshalAs(UnmanagedType.Interface)] [In] IStorage pStgSave,
            [MarshalAs(UnmanagedType.Bool)][In] bool fSameAsLoad);
        [PreserveSig]
        uint SaveCompleted(
            [MarshalAs(UnmanagedType.Interface)] [In] IStorage pStgNew);
        [PreserveSig]
        uint HandsOffStorage();
    }
}
