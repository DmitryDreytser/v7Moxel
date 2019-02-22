using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;

namespace Ole
{
    [Flags]
    public enum STGTY
    {
        STGTY_STORAGE = 1,
        STGTY_STREAM = 2,
        STGTY_LOCKBYTES = 3,
        STGTY_PROPERTY = 4
    }

    [Flags]
    public enum STATFLAG
    {
        STATFLAG_DEFAULT = 0,
        STATFLAG_NONAME = 1,
        STATFLAG_NOOPEN = 2
    }

    [Flags]
    public enum STGC : int
    {
        DEFAULT = 0,
        OVERWRITE = 1,
        ONLYIFCURRENT = 2,
        DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
        CONSOLIDATE = 8
    }

    [Flags]
    public enum STGM : uint
    {
        STGM_READ = 0x00000000,
        STGM_WRITE = 0x00000001,
        STGM_READWRITE = 0x00000002,
        STGM_SHARE_DENY_NONE = 0x00000040,
        STGM_SHARE_DENY_READ = 0x00000030,
        STGM_SHARE_DENY_WRITE = 0x00000020,
        STGM_SHARE_EXCLUSIVE = 0x00000010,
        STGM_PRIORITY = 0x00040000,
        STGM_CREATE = 0x00001000,
        STGM_CONVERT = 0x00020000,
        STGM_FAILIFTHERE = 0x00000000,
        STGM_DIRECT = 0x00000000,
        STGM_TRANSACTED = 0x00010000,
        STGM_NOSCRATCH = 0x00100000,
        STGM_NOSNAPSHOT = 0x00200000,
        STGM_SIMPLE = 0x08000000,
        STGM_DIRECT_SWMR = 0x00400000,
        STGM_DELETEONRELEASE = 0x04000000

    }

    [ComImport]
    [Guid("0000000b-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStorage
    {
        [PreserveSig]
        HRESULT CreateStream(
            /* [string][in] */ string pwcsName,
            /* [in] */ STGM grfMode,
            /* [in] */ uint reserved1,
            /* [in] */ uint reserved2,
            /* [out] */[Out, MarshalAs(UnmanagedType.Interface)] out IStream ppstm);

        [PreserveSig]
        HRESULT OpenStream(
            /* [string][in] */ string pwcsName,
            /* [unique][in] */ IntPtr reserved1,
            /* [in] */ STGM grfMode,
            /* [in] */ uint reserved2,
            /* [out] */[Out, MarshalAs(UnmanagedType.Interface)] out IStream ppstm);

        [PreserveSig]
        HRESULT CreateStorage(
            /* [string][in] */ string pwcsName,
            /* [in] */ STGM grfMode,
            /* [in] */ uint reserved1,
            /* [in] */ uint reserved2,
            /* [out] */[Out, MarshalAs(UnmanagedType.Interface)] out IStorage ppstg);

        [PreserveSig]
        HRESULT OpenStorage(
            /* [string][unique][in] */ string pwcsName,
            /* [unique][in] */ IStorage pstgPriority,
            /* [in] */ STGM grfMode,
            /* [unique][in] */ IntPtr snbExclude,
            /* [in] */ uint reserved,
            /* [out] */[Out, MarshalAs(UnmanagedType.Interface)] out IStorage ppstg);

        [PreserveSig]
        HRESULT CopyTo(
            /* [in] */ uint ciidExclude,
            /* [size_is][unique][in] */ Guid rgiidExclude, // should this be an array?
                                                           /* [unique][in] */ IntPtr snbExclude,
            /* [unique][in] */[In, MarshalAs(UnmanagedType.Interface)] IStorage pstgDest);

        [PreserveSig]
        HRESULT MoveElementTo(
            /* [string][in] */ string pwcsName,
            /* [unique][in] */ [In, MarshalAs(UnmanagedType.Interface)] IStorage pstgDest,
            /* [string][in] */ string pwcsNewName,
            /* [in] */ uint grfFlags);

        [PreserveSig]
        HRESULT Commit(
            /* [in] */ STGC grfCommitFlags);

        [PreserveSig]
        HRESULT Revert();

        [PreserveSig]
        HRESULT EnumElements(
            /* [in] */ uint reserved1,
            /* [size_is][unique][in] */ IntPtr reserved2,
            /* [in] */ uint reserved3,
            /* [out] */[Out, MarshalAs(UnmanagedType.Interface)] out IEnumSTATSTG ppenum);

        [PreserveSig]
        HRESULT DestroyElement(
            /* [string][in] */ string pwcsName);

        [PreserveSig]
        HRESULT RenameElement(
            /* [string][in] */ string pwcsOldName,
            /* [string][in] */ string pwcsNewName);

        [PreserveSig]
        HRESULT SetElementTimes(
            /* [string][unique][in] */ string pwcsName,
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime,
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime,
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

        [PreserveSig]
        HRESULT SetClass(
            /* [in] */ Guid clsid);

        [PreserveSig]
        HRESULT SetStateBits(
            /* [in] */ uint grfStateBits,
            /* [in] */ uint grfMask);

        [PreserveSig]
        HRESULT Stat(
            /* [out] */ out STATSTG pstatstg,
            /* [in] */ STATFLAG grfStatFlag);

    }
}
