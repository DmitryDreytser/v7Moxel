using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Moxel
{
    public class MemoryReader
    {

        public class CSheetDoc : CObject
        {
            MoxelNative.SerializeDelegate _Serialize = MoxelNative.GetSerializer("?Serialize@CSheetDoc@@UAEXAAVCArchive@@@Z"); //CSheetDoc::Serialize(CSheetDoc *this, struct CArchive *Archive)
            public CSheetDoc(IntPtr pMem) : base(pMem)
            {

            }

            public void Serialize(CArchive Arch)
            {
               _Serialize(this, Arch);
            }
        }

        public class CSheet : CObject
        {
            MoxelNative.SerializeDelegate _Serialize = MoxelNative.GetSerializer("?Serialize@CSheet@@UAEXAAVCArchive@@@Z"); //CSheet::Serialize(CSheet *this, CArchive *a2)

            public CSheetDoc SheetDoc = null;

            public CSheet(IntPtr pMem) : base(pMem)
            {
                SheetDoc = GetMember<CSheetDoc>(0x21c);
            }

            public void Serialize(CArchive Arch)
            {
                try
                {
                    _Serialize(pObject, Arch);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        public class CRuntimeClass
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string ClassName;
            public int m_nObjectSize;
            public uint m_wSchema;
            private IntPtr m_pfnCreateObject;
            private IntPtr m_pfnGetBaseClass;

            public object CreateObject()
            {
                if (m_pfnCreateObject != IntPtr.Zero)
                {
                    IntPtr pObj = Marshal.AllocHGlobal(m_nObjectSize);
                    return Marshal.GetDelegateForFunctionPointer<MFCNative._CreateObject>(m_pfnCreateObject).Invoke(pObj);
                }
                else
                    return null;
            }

            public CRuntimeClass GetBaseClass()
            {
                if (m_pfnGetBaseClass != IntPtr.Zero)
                    return Marshal.PtrToStructure<CRuntimeClass>(Marshal.GetDelegateForFunctionPointer<MFCNative._GetBaseClass>(m_pfnGetBaseClass).Invoke());
                else
                    return null;
            }

            public override string ToString()
            {
                return ClassName;
            }

            public CRuntimeClass GetRootClass()
            {
                CRuntimeClass result = null;

                if (m_pfnGetBaseClass != IntPtr.Zero)
                {
                    MFCNative._GetBaseClass fGetbaseclass = Marshal.GetDelegateForFunctionPointer<MFCNative._GetBaseClass>(m_pfnGetBaseClass);
                    if (fGetbaseclass != null)
                    {
                        result = Marshal.PtrToStructure<CRuntimeClass>(fGetbaseclass.Invoke());
                    }
                    if (result == null)
                        return this;
                    else
                        return result.GetRootClass();
                }
                else
                    return null;
            }
        }


        public class CTableOutputContext : CObject
        {
            public CSheet Sheet = null;

            public CTableOutputContext(IntPtr pMem) : base(pMem)
            {
                Sheet = GetMember<CSheet>(56);
            }
        }

        public class CObject
        {
            protected IntPtr pObject;
            public IntPtr Pointer { get { return pObject; } }

            public CRuntimeClass RuntimeClass;

            public byte[] Data
            {
                get
                {
                    byte[] data = new byte[RuntimeClass.m_nObjectSize];
                    Marshal.Copy(pObject, data, 0, data.Length);
                    return data;
                }
            }

            public static implicit operator IntPtr(CObject obj)
            {
                return obj.pObject;
            }

            public static implicit operator CObject(IntPtr pMem)
            {
                return new CObject(pMem);
            }

            public CObject()
            {

            }

            public CObject(IntPtr pMem)
            {
                pObject = pMem;
                IntPtr pGetClass = Marshal.ReadIntPtr(Marshal.ReadIntPtr(pObject, 0), 0);
                byte testByte = Marshal.ReadByte(pGetClass, 0);

                if (testByte == 0xB8) // MOV EAX
                {
                    MFCNative.pGetRuntimeClass GetRuntimeClass = Marshal.GetDelegateForFunctionPointer<MFCNative.pGetRuntimeClass>(pGetClass);
                    RuntimeClass = Marshal.PtrToStructure<CRuntimeClass>(GetRuntimeClass(pObject));
                    if (this.GetType().Name != RuntimeClass.ClassName)
                        throw new Exception($"Получен указатель на объект {RuntimeClass.ClassName} вместо {this.GetType().Name}");
                }
                else
                    throw new Exception($"Указатель не на экземпляр CRuntimeClass ({this.GetType().Name} pGetClass:BYTE [{pGetClass.ToInt32():X8}] = {testByte:X8})");
            }


            public static T FromComObject<T>(object obj) where T : CObject
            {
                IntPtr ptr = Marshal.GetIUnknownForObject(obj); //Возвращает указатель на CBLExportContext
                if (ptr != IntPtr.Zero)
                {
                    IntPtr COutputContext = Marshal.ReadIntPtr(ptr, 8); // укзатель на COutputContext
                    Marshal.Release(ptr);
                    return (T)Activator.CreateInstance(typeof(T), COutputContext);
                }
                else
                    return null;
            }

            public T GetMember<T>(int offset) where T : CObject
            {
                if (pObject != IntPtr.Zero)
                    return (T)Activator.CreateInstance(typeof(T), Marshal.ReadIntPtr(pObject, offset));
                else
                    return null;
            }

            public CObject GetMember(int offset)
            {
                return new CObject(Marshal.ReadIntPtr(pObject, offset));
            }
        }

        public class CFile : CObject
        {
            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            public delegate IntPtr CFileConstructor(IntPtr pMem, IntPtr hFile);

            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            public delegate IntPtr CFileDestructor(IntPtr pMem);

            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            public delegate void CFile__Write(IntPtr _this, IntPtr lpBUf, int nCount);

            private static CFileConstructor _CFile = MFCNative.GetDelegate<CFileConstructor>(352); //??0CFile@@QAE@H@Z CFile::CFile(int HFile)

            private static CFileDestructor _CFileDestructor = MFCNative.GetDelegate<CFileDestructor>(665);

            byte[] buffer = new byte[1024];
            int position = 0;

            public byte[] GetBufer()
            {
                Array.Resize<byte>(ref buffer, position);
                unpatch();
                return buffer;
            }

            public void Write(IntPtr _this, IntPtr lpBUf, int nCount)
            {
                if (buffer.Length < position + nCount + 1)
                    Array.Resize<byte>(ref buffer, (position + nCount) * 2);

                Marshal.Copy(lpBUf, buffer, position, nCount);
                position += nCount;
            }

            public CFile__Write WriteDelegate = null;
            static IntPtr pWriteDelegate;
            static GCHandle hh;
            static IntPtr pVTable = IntPtr.Zero;
            static IntPtr old_Func;
            static IntPtr FuncAddr;
            static bool patched = false;

            void patch()
            {
                if (patched)
                    return;

                Handle = new CFile__Write(Write);
                hh = GCHandle.Alloc(Handle);

                pWriteDelegate = Marshal.GetFunctionPointerForDelegate<CFile__Write>(Handle);

                uint OldProtection;

                WinApi.VirtualProtectEx(Process.GetCurrentProcess().Handle, FuncAddr, new IntPtr(4), 0x40, out OldProtection);

                Marshal.WriteIntPtr(FuncAddr, pWriteDelegate);

                WinApi.VirtualProtectEx(Process.GetCurrentProcess().Handle, FuncAddr, new IntPtr(4), OldProtection, out OldProtection);

                patched = true;
            }

            static CFile__Write Handle;

            void unpatch()
            {
                uint OldProtection;
                WinApi.VirtualProtectEx(Process.GetCurrentProcess().Handle, FuncAddr, new IntPtr(4), 0x40, out OldProtection);
                Marshal.WriteIntPtr(FuncAddr, old_Func); // Вернем оригинальную функцию.
                WinApi.VirtualProtectEx(Process.GetCurrentProcess().Handle, FuncAddr, new IntPtr(4), OldProtection, out OldProtection);
                hh.Free();

                patched = false;
            }

            public CFile(IntPtr pMem) : base(pMem)
            {
                // перехватим CFile::Write()
                pVTable = Marshal.ReadIntPtr(pMem, 0);
                old_Func = Marshal.ReadIntPtr(pVTable, 0x40); // CFile::Write находитс по смещение 0x40 от начала таблицы виртуальных функций.
                FuncAddr = new IntPtr((Int64)pVTable + 0x40);
                patch();
            }

            public static CFile FromHFile(IntPtr hFile)
            {
                IntPtr pMem = Marshal.AllocHGlobal(0x10);
                return new CFile(_CFile(pMem, hFile));
            }

            ~CFile()
            {
                _CFileDestructor(this);
                Marshal.FreeHGlobal(pObject);
            }

        }

        public class CArchive
        {
            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            public delegate void CArchive_Flush(IntPtr _this);

            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            delegate IntPtr CArchiveConstructor(IntPtr pObject, IntPtr CFile, Mode nMode, int nBufSize, IntPtr lpBuf);

            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            delegate IntPtr CArchiveDestructor(IntPtr pObject);

            static CArchive_Flush flush = MFCNative.GetDelegate<CArchive_Flush>(2801);
            static CArchiveConstructor _CArchive = MFCNative.GetDelegate<CArchiveConstructor>(273); //??0CArchive@@QAE@PAVCFile@@IHPAX@Z CArchive::CArchive(CFile* pFile, UINT nMode, int nBufSize, void* lpBuf)
            static CArchiveDestructor _CArchiveDestructor = MFCNative.GetDelegate<CArchiveDestructor>(603);

            //[StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Ansi)]
            //struct CArchiveInternal
            //{
            //    IntPtr m_pDocument;
            //    [MarshalAs( UnmanagedType.Bool)]
            //    bool m_bForceFlat;
            //    [MarshalAs(UnmanagedType.Bool)]
            //    bool m_bDirectBuffer;
            //    uint m_nObjectSchema;
            //    [MarshalAs(UnmanagedType.LPStr)]
            //    string m_strFileName;
            //    [MarshalAs(UnmanagedType.Bool)]
            //    bool m_nMode;
            //    [MarshalAs(UnmanagedType.Bool)]
            //    bool m_bUserBuf;
            //    int m_nBufSize;
            //    IntPtr m_pFile;
            //    IntPtr m_lpBufCur;
            //    IntPtr m_lpBufMax;
            //    IntPtr m_lpBufStart;
            //    uint m_nMapCount;
            //    IntPtr m_pLoadArray;
            //    IntPtr m_pSchemaMap;
            //    uint m_nGrowSize;
            //    uint m_nHashSize;
            //};

            // CArchiveInternal ArchiveStruct;

            public IntPtr pDocument
            {
                set
                {
                    Marshal.WriteIntPtr(pObject, value); // CDocument по смещению 0 от начала объекта в памяти
                }
            }

            public IntPtr pObject;
            enum Mode : uint { store = 0, load = 1, bNoFlushOnDelete = 2, bNoByteSwap = 4 };


            public CArchive(CFile pfile, IntPtr CDocument)
            {
                pObject = Marshal.AllocHGlobal(0x44);
                pObject = _CArchive(pObject, pfile, Mode.store, 1024 * 4096, IntPtr.Zero);
                pDocument = CDocument;
                //ArchiveStruct = Marshal.PtrToStructure<CArchiveInternal>(pObject);
            }

            public void Flush()
            {
                flush(this);
            }

            public static implicit operator IntPtr(CArchive obj)
            {
                return obj.pObject;
            }

            ~CArchive()
            {
                _CArchiveDestructor(this);
                Marshal.FreeHGlobal(this);
            }
        }
    }

}
