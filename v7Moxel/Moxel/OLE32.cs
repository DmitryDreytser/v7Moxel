using Ole;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace Moxel
{
    public static class OLE32
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct STGOPTIONS
        {
            ushort usVersion;
            ushort reserved;
            ulong ulSectorSize;
            [MarshalAsAttribute(UnmanagedType.LPWStr)]
            public string pwcsTemplateFile;
        };

        public enum OLERENDER
        {
            OLERENDER_NONE = 0,
            OLERENDER_DRAW = 1,
            OLERENDER_FORMAT = 2,
            OLERENDER_ASIS = 3
        }


        internal enum CLIPFORMAT : int
        {
            CF_TEXT = 1,
            CF_BITMAP = 2,
            CF_METAFILEPICT = 3,
            CF_SYLK = 4,
            CF_DIF = 5,
            CF_TIFF = 6,
            CF_OEMTEXT = 7,
            CF_DIB = 8,
            CF_PALETTE = 9,
            CF_PENDATA = 10,
            CF_RIFF = 11,
            CF_WAVE = 12,
            CF_UNICODETEXT = 13,
            CF_ENHMETAFILE = 14,
            CF_HDROP = 15,
            CF_LOCALE = 16,
            CF_MAX = 17,
            CF_OWNERDISPLAY = 0x80,
            CF_DSPTEXT = 0x81,
            CF_DSPBITMAP = 0x82,
            CF_DSPMETAFILEPICT = 0x83,
            CF_DSPENHMETAFILE = 0x8E,
        }

        internal enum STGFMT : int
        {
            STGFMT_STORAGE = 0,
            STGFMT_FILE = 3,
            STGFMT_ANY = 4,
            STGFMT_DOCFILE = 5
        }

        [Flags]
        public  enum STGM : int
        {
            STGM_READ = 0x0,
            STGM_WRITE = 0x1,
            STGM_READWRITE = 0x2,
            STGM_SHARE_DENY_NONE = 0x40,
            STGM_SHARE_DENY_READ = 0x30,
            STGM_SHARE_DENY_WRITE = 0x20,
            STGM_SHARE_EXCLUSIVE = 0x10,
            STGM_PRIORITY = 0x40000,
            STGM_CREATE = 0x1000,
            STGM_CONVERT = 0x20000,
            STGM_FAILIFTHERE = 0x0,
            STGM_DIRECT = 0x0,
            STGM_TRANSACTED = 0x10000,
            STGM_NOSCRATCH = 0x100000,
            STGM_NOSNAPSHOT = 0x200000,
            STGM_SIMPLE = 0x8000000,
            STGM_DIRECT_SWMR = 0x400000,
            STGM_DELETEONRELEASE = 0x4000000
        }

        public enum STGC
        {
            DEFAULT = 0,
            OVERWRITE = 1,
            ONLYIFCURRENT = 2,
            DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
            CONSOLIDATE = 8
        }

        public enum CoInit
        {
            MultiThreaded = 0x0,
            ApartmentThreaded = 0x2,
            DisableOLE1DDE = 0x4,
            SpeedOverMemory = 0x8
        }


        public static Guid IID_IDataObject = new Guid("{0000010e-0000-0000-C000-000000000046}");

        public static Guid IID_IOleObject = new Guid("{00000112-0000-0000-C000-000000000046}");

        public static Guid IID_IStorage = new Guid("0000000B-0000-0000-C000-000000000046");



        [DllImport("ole32.dll")]
        public static extern int StgCreateStorageEx([MarshalAs(UnmanagedType.LPWStr)] string
           pwcsName, int grfMode, int stgfmt, int grfAttrs, IntPtr pStgOptions, IntPtr reserved2, ref Guid riid,
           out IStorage ppObjectOpen);


        [DllImport("ole32.dll", PreserveSig = false)]
        [return: MarshalAs(UnmanagedType.IUnknown)]
        public static extern object OleGetClipboard();

        [DllImport("ole32.dll")]
        public static extern int OleCreateFromFile([In] ref Guid rclsid,
           [MarshalAs(UnmanagedType.LPWStr)] string lpszFileName, ref Guid riid,
           uint renderopt, IntPtr pFormatEtc, IOleClientSite pClientSite,
           IStorage pStg, out IOleObject ppvObj);



        [DllImport("ole32.dll")]
        public static extern int OleCreateFromFile([In] ref Guid rclsid,
           [MarshalAs(UnmanagedType.LPWStr)] string lpszFileName, [In] ref Guid riid,
           uint renderopt, [In] ref FORMATETC pFormatEtc, IOleClientSite pClientSite,
           IStorage pStg, out IOleObject ppvObj);

        [DllImport("ole32.dll")]
        public static extern int OleCreateFromFileEx([In] ref Guid rclsid,
           [MarshalAs(UnmanagedType.LPWStr)] string lpszFileName, [In] ref Guid riid,
           uint dwFlags, uint renderopt, uint cFormats, uint rgAdvf, FORMATETC[]
           rgFormatEtc, IAdviseSink pAdviseSink, [Out] uint[] rgdwConnection,
           IOleClientSite pClientSite, IStorage pStg,
           [MarshalAs(UnmanagedType.IUnknown)] out object ppvObj);


        [DllImport("ole32.dll")]
        public static extern int OleSetContainedObject([MarshalAs(UnmanagedType.IUnknown)] object pUnk, bool fContained);

        [DllImport("ole32.dll")]
        public static extern Ole.HRESULT OleRun([MarshalAs(UnmanagedType.IUnknown)] object pUnknown);

        [DllImport("Ole32.dll", ExactSpelling = true, EntryPoint = "CoInitialize",
   CallingConvention = CallingConvention.StdCall, SetLastError = false)]
        public static extern int CoInitialize(
            IntPtr pvReserved
            );

        [DllImport("Ole32.dll", ExactSpelling = true, EntryPoint = "CoInitializeEx", CallingConvention = CallingConvention.StdCall, SetLastError = false, PreserveSig = false)]
        public static extern void CoInitializeEx(IntPtr pvReserved, CoInit coInit);

        [DllImport("Ole32.dll", ExactSpelling = true, EntryPoint = "CoUninitialize",
   CallingConvention = CallingConvention.StdCall, SetLastError = false)]
        public static extern int CoUninitialize();


        [DllImport("Ole32.dll", ExactSpelling = true, EntryPoint = "CoInitializeSecurity",
       CallingConvention = CallingConvention.StdCall, SetLastError = false)]
        static extern int CoInitializeSecurity(
            IntPtr pVoid,
            uint cAuthSvc,
            IntPtr asAuthSvc,
            IntPtr pReserved1,
            uint dwAuthnLevel,
            uint dwImpLevel,
            IntPtr pAuthList,
            uint dwCapabilities,
            IntPtr pReserved3
            );




        [DllImport("ole32.dll")]
        public extern static int CreateILockBytesOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool fDeleteOnRelease, out ILockBytes ppLkbyt);

        [DllImport("ole32.dll")]
        public extern static int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, StgmConstants grfMode, int reserved, out IStorage ppstgOpen);

        [DllImport("ole32.dll")]
        [PreserveSig()]
        public static extern Ole.HRESULT StgOpenStorageOnILockBytes(ILockBytes plkbyt,
         IStorage pStgPriority, STGM grfMode, IntPtr snbEnclude, uint reserved,
         out IStorage ppstgOpen);

        [DllImport("ole32.dll")]
        public static extern Ole.HRESULT OleCreateFromFile([In] ref Guid rclsid,
           [MarshalAs(UnmanagedType.LPWStr)] string lpszFileName, [In] ref Guid riid,
           uint renderopt, [In] ref FORMATETC pFormatEtc, IOleClientSite pClientSite,
           IStorage pStg, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObj);


        [DllImport("ole32.dll")]
        public static extern int CreateStreamOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease,
           out System.Runtime.InteropServices.ComTypes.IStream ppstm);

        [DllImport("ole32.dll")]
        public static extern Ole.HRESULT OleLoadFromStream(
           System.Runtime.InteropServices.ComTypes.IStream pStm,
           [In] ref Guid riid,
           out IOleObject ppvObj);

        [DllImport("ole32.dll", SetLastError = true)]
        public static extern Ole.HRESULT OleLoad(
          [In] IStorage       pStg,
          [In] ref Guid           riid,
          [In] IOleClientSite pClientSite,
          [Out] out IOleObject  ppvObj
            );

        [DllImport("ole32.dll")]
        public static extern Ole.HRESULT OleDraw([MarshalAs(UnmanagedType.IUnknown)] object pUnk, uint dwAspect, HandleRef hdcDraw, ref Rectangle lprcBounds);

        [DllImport("ole32.dll")]
        public static extern Ole.HRESULT OleInitialize(IntPtr rezerved);




        [Flags]
        public enum StgmConstants
        {
            STGM_READ = 0x0,
            STGM_WRITE = 0x1,
            STGM_READWRITE = 0x2,
            STGM_SHARE_DENY_NONE = 0x40,
            STGM_SHARE_DENY_READ = 0x30,
            STGM_SHARE_DENY_WRITE = 0x20,
            STGM_SHARE_EXCLUSIVE = 0x10,
            STGM_PRIORITY = 0x40000,
            STGM_CREATE = 0x1000,
            STGM_CONVERT = 0x20000,
            STGM_FAILIFTHERE = 0x0,
            STGM_DIRECT = 0x0,
            STGM_TRANSACTED = 0x10000,
            STGM_NOSCRATCH = 0x100000,
            STGM_NOSNAPSHOT = 0x200000,
            STGM_SIMPLE = 0x8000000,
            STGM_DIRECT_SWMR = 0x400000,
            STGM_DELETEONRELEASE = 0x4000000
        }





        [ComVisible(false)]
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000A-0000-0000-C000-000000000046")]
        public interface ILockBytes
        {
            void ReadAt(long ulOffset, IntPtr pv, int cb, out UIntPtr pcbRead);
            void WriteAt(long ulOffset, IntPtr pv, int cb, out UIntPtr pcbWritten);
            void Flush();
            void SetSize(long cb);
            void LockRegion(long libOffset, long cb, int dwLockType);
            void UnlockRegion(long libOffset, long cb, int dwLockType);
            void Stat(out System.Runtime.InteropServices.STATSTG pstatstg, int grfStatFlag);

        }

        [ComImport]
        [Guid("0000000b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IStorage
        {
            void CreateStream(
                /* [string][in] */ string pwcsName,
                /* [in] */ StgmConstants grfMode,
                /* [in] */ uint reserved1,
                /* [in] */ uint reserved2,
                /* [out] */ out IStream ppstm);

            void OpenStream(
                /* [string][in] */ string pwcsName,
                /* [unique][in] */ IntPtr reserved1,
                /* [in] */ uint grfMode,
                /* [in] */ uint reserved2,
                /* [out] */ out IStream ppstm);

            void CreateStorage(
                /* [string][in] */ string pwcsName,
                /* [in] */ uint grfMode,
                /* [in] */ uint reserved1,
                /* [in] */ uint reserved2,
                /* [out] */ out IStorage ppstg);

            void OpenStorage(
                /* [string][unique][in] */ string pwcsName,
                /* [unique][in] */ IStorage pstgPriority,
                /* [in] */ uint grfMode,
                /* [unique][in] */ IntPtr snbExclude,
                /* [in] */ uint reserved,
                /* [out] */ out IStorage ppstg);

            void CopyTo(
                /* [in] */ uint ciidExclude,
                /* [size_is][unique][in] */ Guid rgiidExclude, // should this be an array?
                                                               /* [unique][in] */ IntPtr snbExclude,
                /* [unique][in] */ IStorage pstgDest);

            void MoveElementTo(
                /* [string][in] */ string pwcsName,
                /* [unique][in] */ IStorage pstgDest,
                /* [string][in] */ string pwcsNewName,
                /* [in] */ uint grfFlags);

            void Commit(
                /* [in] */ uint grfCommitFlags);

            void Revert();

            void EnumElements(
                /* [in] */ uint reserved1,
                /* [size_is][unique][in] */ IntPtr reserved2,
                /* [in] */ uint reserved3,
                /* [out] */ out IEnumSTATSTG ppenum);

            void DestroyElement(
                /* [string][in] */ string pwcsName);

            void RenameElement(
                /* [string][in] */ string pwcsOldName,
                /* [string][in] */ string pwcsNewName);

            void SetElementTimes(
                /* [string][unique][in] */ string pwcsName,
                /* [unique][in] */ System.Runtime.InteropServices.FILETIME pctime,
                /* [unique][in] */ System.Runtime.InteropServices.FILETIME patime,
                /* [unique][in] */ System.Runtime.InteropServices.FILETIME pmtime);

            void SetClass(
                /* [in] */ Guid clsid);

            void SetStateBits(
                /* [in] */ uint grfStateBits,
                /* [in] */ uint grfMask);

            void Stat(
                /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
                /* [in] */ uint grfStatFlag);

        }

        [ComImport]
        [Guid("00000112-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleObject
        {
            void SetClientSite(IOleClientSite pClientSite);
            void GetClientSite(ref IOleClientSite ppClientSite);
            void SetHostNames(object szContainerApp, object szContainerObj);
            void Close(uint dwSaveOption);
            void SetMoniker(uint dwWhichMoniker, object pmk);
            void GetMoniker(uint dwAssign, uint dwWhichMoniker, object ppmk);
            void InitFromData(IDataObject pDataObject, bool fCreation, uint dwReserved);
            void GetClipboardData(uint dwReserved, ref IDataObject ppDataObject);
            void DoVerb(uint iVerb, uint lpmsg, object pActiveSite, uint lindex, uint hwndParent, uint lprcPosRect);
            void EnumVerbs(ref object ppEnumOleVerb);
            void Update();
            void IsUpToDate();
            void GetUserClassID(uint pClsid);
            void GetUserType(uint dwFormOfType, uint pszUserType);
            void SetExtent(uint dwDrawAspect, [In] ref Ole.tagSIZEL psizel);
            void GetExtent(uint dwDrawAspect, [In] ref Ole.tagSIZEL psizel);
            void Advise(object pAdvSink, uint pdwConnection);
            void Unadvise(uint dwConnection);
            void EnumAdvise(ref object ppenumAdvise);
            void GetMiscStatus(uint dwAspect, uint pdwStatus);
            void SetColorScheme(object pLogpal);
        };

        [ComImport]
        [Guid("0000000d-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumSTATSTG
        {
            // The user needs to allocate an STATSTG array whose size is celt.
            [PreserveSig]
            uint
            Next(
            uint celt,
            [MarshalAs(UnmanagedType.LPArray), Out]
    System.Runtime.InteropServices.STATSTG[] rgelt,
            out uint pceltFetched
            );

            void Skip(uint celt);

            void Reset();

            [return: MarshalAs(UnmanagedType.Interface)]
            IEnumSTATSTG Clone();
        }


        [ComImport, Guid("0000000c-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IStream
        {
            void Read([Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, uint cb, out uint pcbRead);
            void Write([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, uint cb, out uint pcbWritten);
            void Seek(long dlibMove, uint dwOrigin, out long plibNewPosition);
            void SetSize(long libNewSize);
            void CopyTo(IStream pstm, long cb, out long pcbRead, out long pcbWritten);
            void Commit(uint grfCommitFlags);
            void Revert();
            void LockRegion(long libOffset, long cb, uint dwLockType);
            void UnlockRegion(long libOffset, long cb, uint dwLockType);
            void Stat(out System.Runtime.InteropServices.STATSTG pstatstg, uint grfStatFlag);
            void Clone(out IStream ppstm);
        }


        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000010E-0000-0000-C000-000000000046")]
        public interface IDataObject
        {
            void GetData([In] ref FORMATETC format, out STGMEDIUM medium);
            void GetDataHere([In] ref FORMATETC format, ref STGMEDIUM medium);
            [PreserveSig]
            int QueryGetData([In] ref FORMATETC format);
            [PreserveSig]
            int GetCanonicalFormatEtc([In] ref FORMATETC formatIn, out FORMATETC formatOut);
            void SetData([In] ref FORMATETC formatIn, [In] ref STGMEDIUM medium, [MarshalAs(UnmanagedType.Bool)] bool release);
            IEnumFORMATETC EnumFormatEtc(DATADIR direction);
            [PreserveSig]
            int DAdvise([In] ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection);
            void DUnadvise(int connection);
            [PreserveSig]
            int EnumDAdvise(out IEnumSTATDATA enumAdvise);
        }
    }

    public class Helper
    {
        public static string ExportOleFile(string inputFileName, string oleOutputFileName, string emfOutputFileName)
        {
            StringBuilder resultString = new StringBuilder();
            try
            {
                OLE32.CoUninitialize();
                OLE32.CoInitializeEx(IntPtr.Zero, OLE32.CoInit.ApartmentThreaded); //COINIT_APARTMENTTHREADED

                OLE32.IStorage storage;
                var result = OLE32.StgCreateStorageEx(oleOutputFileName,
                    Convert.ToInt32(OLE32.STGM.STGM_READWRITE | OLE32.STGM.STGM_SHARE_EXCLUSIVE | OLE32.STGM.STGM_CREATE | OLE32.STGM.STGM_TRANSACTED),
                    Convert.ToInt32(OLE32.STGFMT.STGFMT_DOCFILE),
                    0,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    ref OLE32.IID_IStorage,
                    out storage
                );

                if (result != 0) result.ToString("X");

                var CLSID_NULL = Guid.Empty;

                OLE32.IOleObject pOle;
                result = OLE32.OleCreateFromFile(
                    ref CLSID_NULL,
                    inputFileName,
                    ref OLE32.IID_IDataObject,
                    (uint)OLE32.OLERENDER.OLERENDER_NONE,
                    IntPtr.Zero,
                    null,
                    storage,
                    out pOle
                );

                if (result != 0) return result.ToString("X");



                result = (int)OLE32.OleRun(pOle);
                

            }
            catch (Exception ex)
            {
                resultString.AppendLine(ex.ToString());
                return resultString.ToString();
            }
            return resultString.ToString();
        }
    }
}
