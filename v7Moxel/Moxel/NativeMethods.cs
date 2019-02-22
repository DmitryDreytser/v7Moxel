using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace Moxel
{
    static class NativeMethods
    {
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
                    System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt,
                out uint pceltFetched
            );
            void Skip(uint celt);
            void Reset();
            [return: MarshalAs(UnmanagedType.Interface)]
            IEnumSTATSTG Clone();
        }

        [ComImport]
        [Guid("0000000b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IStorage
        {
            void CreateStream(
                /* [string][in]*/  string pwcsName,
                /* [in]*/  STGM grfMode,
                /* [in]*/  uint reserved1,
                /* [in] */ uint reserved2,
                /* [out] */ out IStream ppstm);

            void OpenStream(
                /* [string][in] */ string pwcsName,
                /* [unique][in] */ IntPtr reserved1,
                /* [in] */ STGM grfMode,
                /* [in] */ uint reserved2,
                /* [out] */ out IStream ppstm);

            void CreateStorage(
                /* [string][in] */ string pwcsName,
                /* [in] */ STGM grfMode,
                /* [in] */ uint reserved1,
                /* [in] */ uint reserved2,
                /* [out] */ out IStorage ppstg);

            void OpenStorage(
              /* [string][unique][in] */ string pwcsName,
              /* [unique][in] */ IStorage pstgPriority,
              /* [in] */ STGM grfMode,
              /* [unique][in] */ IntPtr snbExclude,
              /* [in] */ uint reserved,
              /* [out] */ out IStorage ppstg);

            void CopyTo(
               /* [in] */ uint ciidExclude,
               /* [size_is][unique][in] */ ref Guid rgiidExclude, // should this be an array?
                                                                  /* [unique][in] */ IntPtr snbExclude,
               /* [unique][in] */ IStorage pstgDest);

            void MoveElementTo(
               /* [string][in] */ string pwcsName,
               /* [unique][in] */ IStorage pstgDest,
               /* [string][in] */ string pwcsNewName,
               /* [in] */ uint grfFlags);

            void Commit(
               /* [in] */ STGC grfCommitFlags);

            void Revert();

            void EnumElements(
               /* [in] */ uint reserved1,
               /* [size_is][unique][in] */ IntPtr reserved2,
               /* [in] */ uint reserved3,
               /* [out] */out IEnumSTATSTG ppenum);

            void DestroyElement(
               /* [string][in] */ string pwcsName);

            void RenameElement(
               /* [string][in] */ string pwcsOldName,
               /* [string][in] */ string pwcsNewName);

            void SetElementTimes(
               /* [string][unique][in] */ string pwcsName,
               /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime,
               /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime,
               /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

            void SetClass(
               /* [in] */ ref Guid clsid);

            void SetStateBits(
               /* [in] */ uint grfStateBits,
               /* [in] */ uint grfMask);

            void Stat(
               /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
               /* [in] */ uint grfStatFlag);

            //////////////////////////////////////////////////////////////////////////////////


        }

        public class Storage
        {
            private Dictionary<string, NativeMethods.IStorage> STGRefs = new Dictionary<string, NativeMethods.IStorage>();
            public NativeMethods.IStorage RootStorage;

            public Storage(IStorage RootStorage)
            {
                this.RootStorage = RootStorage;
                STGRefs["ROOT"] = this.RootStorage;
            }

            public NativeMethods.IStorage GetStorage(string Path)
            {
                if (STGRefs.ContainsKey(Path.ToUpper()))
                    return STGRefs[Path.ToUpper()];

                NativeMethods.IStorage newStorage;
                string[] chunks = Path.Split('\\');
                string name = chunks.Last();
                string parentPath = Path.Replace("\\" + name, "");

                NativeMethods.IStorage parentStorage = GetStorage(parentPath);
                parentStorage.OpenStorage(name, null, NativeMethods.STGM.READWRITE | NativeMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0, out newStorage);
                STGRefs[Path.ToUpper()] = newStorage;
                return newStorage;
            }

            ~Storage()
            {
                foreach (IStorage item in STGRefs.Values)
                {
                    Marshal.ReleaseComObject(item);
                    Marshal.FinalReleaseComObject(item);
                }

                STGRefs.Clear();
            }
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
        public enum STGM : int
        {
            DIRECT = 0x00000000,
            TRANSACTED = 0x00010000,
            SIMPLE = 0x08000000,
            READ = 0x00000000,
            WRITE = 0x00000001,
            READWRITE = 0x00000002,
            SHARE_DENY_NONE = 0x00000040,
            SHARE_DENY_READ = 0x00000030,
            SHARE_DENY_WRITE = 0x00000020,
            SHARE_EXCLUSIVE = 0x00000010,
            PRIORITY = 0x00040000,
            DELETEONRELEASE = 0x04000000,
            NOSCRATCH = 0x00100000,
            CREATE = 0x00001000,
            CONVERT = 0x00020000,
            FAILIFTHERE = 0x00000000,
            NOSNAPSHOT = 0x00200000,
            DIRECT_SWMR = 0x00400000,
        }

        [Flags]
        public enum STATFLAG : uint
        {
            STATFLAG_DEFAULT = 0,
            STATFLAG_NONAME = 1,
            STATFLAG_NOOPEN = 2
        }

        [Flags]
        public enum STGTY : int
        {
            STGTY_STORAGE = 1,
            STGTY_STREAM = 2,
            STGTY_LOCKBYTES = 3,
            STGTY_PROPERTY = 4
        }

        //Читает IStream в массив байт
        public static byte[] ReadIStream(IStream pIStream)
        {
            System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
            pIStream.Stat(out StreamInfo, 0);
            byte[] data = new byte[StreamInfo.cbSize];
            pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);
            Marshal.ReleaseComObject(pIStream);
            Marshal.FinalReleaseComObject(pIStream);
            pIStream = null;
            return data;
        }

        //Читает IStream в массив байт и разжимает
        public static byte[] ReadCompressedIStream(ref IStream pIStream)
        {
            System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
            pIStream.Stat(out StreamInfo, 0);
            byte[] data = new byte[StreamInfo.cbSize];
            pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);

            Marshal.ReleaseComObject(pIStream);
            Marshal.FinalReleaseComObject(pIStream);
            pIStream = null;

            DeflateStream ZLibCompressed = new DeflateStream(new MemoryStream(data), CompressionMode.Decompress, false);
            MemoryStream Decompressed = new MemoryStream();
            ZLibCompressed.CopyTo(Decompressed);
            ZLibCompressed.Dispose();
            data = Decompressed.ToArray();
            Array.Resize(ref data, (int)Decompressed.Length);


            return data;
        }

        //Читает IStream в строку
        public static string ReadIStreamToString(ref IStream pIStream)
        {
            System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
            pIStream.Stat(out StreamInfo, 0);
            byte[] data = new byte[StreamInfo.cbSize];
            pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);
            Marshal.ReleaseComObject(pIStream);
            Marshal.FinalReleaseComObject(pIStream);
            pIStream = null;

            return Encoding.GetEncoding(1251).GetString(data);
        }

        [ComVisible(false)]
        [ComConversionLossAttribute]
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("0000000A-0000-0000-C000-000000000046")]
        public interface ILockBytes
        {
            void ReadAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbRead);
            void WriteAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbWritten);
            void Flush();
            void SetSize(long cb);
            void LockRegion(long libOffset, long cb, int dwLockType);
            void UnlockRegion(long libOffset, long cb, int dwLockType);
            void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag);
        }


        [DllImport("ole32.dll")]
        public static extern int StgIsStorageFile(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

        [DllImport("ole32.dll")]
        public static extern int StgOpenStorage(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
            IStorage pstgPriority,
            STGM grfMode,
            IntPtr snbExclude,
            uint reserved,
            out IStorage ppstgOpen);

        [DllImport("ole32.dll")]
        public static extern int StgCreateDocfile(
            [MarshalAs(UnmanagedType.LPWStr)]string pwcsName,
            STGM grfMode,
            uint reserved,
            out IStorage ppstgOpen);

        [DllImport("ole32.dll")]
        [PreserveSig()]
        public static extern int StgOpenStorageOnILockBytes(ILockBytes plkbyt,
           IStorage pStgPriority, STGM grfMode, IntPtr snbEnclude, uint reserved,
           out IStorage ppstgOpen);

        [DllImport("ole32.dll")]
        public static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool fDeleteOnRelease, out ILockBytes ppLkbyt);

        [DllImport("ole32.dll")]
        public static extern int OleLoadFromStream(IStream pStm,
           [In] ref Guid riid,
           [MarshalAs(UnmanagedType.IUnknown)] out object ppvObj);

        [DllImport("ole32.dll")]
        public static extern int OleLoad(IStorage pStg,
           [In] ref Guid riid,
           [MarshalAs(UnmanagedType.IUnknown)] out object ppvObj);


    }
}
