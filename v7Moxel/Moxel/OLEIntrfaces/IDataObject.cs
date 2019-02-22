using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;

namespace Ole



{
    public enum TYMED
    {
        TYMED_HGLOBAL = 1,
        TYMED_FILE = 2,
        TYMED_ISTREAM = 4,
        TYMED_ISTORAGE = 8,
        TYMED_GDI = 16,
        TYMED_MFPICT = 32,
        TYMED_ENHMF = 64,
        TYMED_NULL = 0
    }
    

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct StgMedium
    {
        public TYMED tymed;
        public IntPtr unionmember;
        [MarshalAs(UnmanagedType.IUnknown)]
        public object pUnkForRelease;
    }

    public enum ClipBoardFormats :ushort
    {
        CF_BITMAP = 2,
        CF_DIB = 8,
        CF_DIBV5 = 17,
        CF_DIF = 5,
        CF_DSPBITMAP = 0x0082,
        CF_DSPENHMETAFILE = 0x008E,
        CF_DSPMETAFILEPICT = 0x0083,
        CF_DSPTEXT = 0x0081,
        CF_ENHMETAFILE = 14,
        CF_GDIOBJFIRST = 0x0300,
        CF_GDIOBJLAST = 0x03FF,
        CF_HDROP = 15,
        CF_LOCALE = 16,
        CF_METAFILEPICT = 3,
        CF_OEMTEXT = 7,
        CF_OWNERDISPLAY = 0x0080,
        CF_PALETTE = 9,
        CF_PENDATA = 10,
        CF_PRIVATEFIRST = 0x0200,
        CF_PRIVATELAST = 0x02FF,
        CF_RIFF = 11,
        CF_SYLK = 4,
        CF_TEXT = 1,
        CF_TIFF = 6,
        CF_UNICODETEXT = 13,
        CF_WAVE = 12
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct FormatEtc
    {
        public ClipBoardFormats cfFormat;
        public IntPtr ptd;
        public uint dwAspect;
        public int lindex;
        public TYMED tymed;
    }

    [Guid("0000010E-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IDataObject
    {
        [PreserveSig]
        int GetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] ref FormatEtc pformatetcIn,
            [MarshalAs(UnmanagedType.LPArray)] [In] ref StgMedium pRemoteMedium);
        [PreserveSig]
        int GetDataHere(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [MarshalAs(UnmanagedType.LPArray)] [In] [Out] StgMedium[] pRemoteMedium);
        [PreserveSig]
        int QueryGetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc);
        [PreserveSig]
        int GetCanonicalFormatEtc(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pformatectIn,
            [MarshalAs(UnmanagedType.LPArray)] [Out] FormatEtc[] pformatetcOut);
        [PreserveSig]
        int SetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [MarshalAs(UnmanagedType.LPArray)] [In] StgMedium[] pmedium,
            [In] int fRelease);
        [PreserveSig]
        int EnumFormatEtc(
            [In] uint dwDirection,
            [MarshalAs(UnmanagedType.Interface)] out IEnumFormatEtc ppenumFormatEtc);
        [PreserveSig]
        int DAdvise(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [In] uint ADVF,
            [MarshalAs(UnmanagedType.Interface)] [In] IAdviseSink pAdvSink,
            out uint pdwConnection);
        [PreserveSig]
        int DUnadvise(
            [In] uint dwConnection);
        [PreserveSig]
        int EnumDAdvise(
            [MarshalAs(UnmanagedType.Interface)] out IEnumStatData ppenumAdvise);
    }
}
