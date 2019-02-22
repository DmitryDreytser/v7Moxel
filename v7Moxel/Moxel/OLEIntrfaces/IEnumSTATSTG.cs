using System;
using System.Runtime.InteropServices;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;

namespace Ole
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct StatData
    {
        public FormatEtc FORMATETC;
        public uint ADVF;
        [MarshalAs(UnmanagedType.Interface)]
        public IAdviseSink pAdvSink;
        public uint dwConnection;
    }
    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct STATSTG
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pwcsName;
        public STGTY type;
        public ulong cbSize;
        public System.Runtime.InteropServices.ComTypes.FILETIME mtime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ctime;
        public System.Runtime.InteropServices.ComTypes.FILETIME atime;
        public STGM grfMode;
        public uint grfLocksSupported;
        public Guid clsid;
        public STATFLAG grfStateBits;
        public uint reserved;
    }

    [ComImport]
    [Guid("0000000d-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumSTATSTG
    {
        // The user needs to allocate an STATSTG array whose size is celt.
        [PreserveSig]
        uint Next(
            uint celt,
            [MarshalAs(UnmanagedType.LPArray), Out]
                        STATSTG[] rgelt,
            out uint pceltFetched
        );
        void Skip(uint celt);
        void Reset();
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumSTATSTG Clone();
    }
}
