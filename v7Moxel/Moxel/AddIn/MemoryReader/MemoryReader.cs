using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Moxel
{
    public class MemoryReader
    {
        public static byte[] ReadMoxel(IntPtr pSheetDoc)
        {
            CSheetDoc SheetDoc = new CSheetDoc(pSheetDoc);
            using (CFile f = CFile.Create(SheetDoc.Length))
            {
                try
                {
                    using (CArchive Arch = new CArchive(f, SheetDoc))
                    {
                        SheetDoc.Serialize(Arch);
                        return f.GetBufer();
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message, ex);
                }
            }
        }

        public static Moxel ReadFromCSheetDoc(CSheetDoc SheetDoc)
        {
            //var pcStream = new Nito.ProducerConsumerStream.ProducerConsumerStream();

            using (var ms = new FileStream(Path.GetTempFileName(), FileMode.Create, FileAccess.ReadWrite))
            {
                using (CFile f = CFile.FromStream(ms))
            {
                try
                {
                    using (CArchive Arch = new CArchive(f, SheetDoc))
                    {
                        SheetDoc.Serialize(Arch);
                        //await Task.Factory.StartNew(() => SheetDoc.Serialize(Arch), TaskCreationOptions.LongRunning);
                        Arch.Flush();
                        return new Moxel(ms);
                        //return await Task.Factory.StartNew(() => new Moxel(ms), TaskCreationOptions.LongRunning);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message, ex);
                }
            }
            }
        }

        public static Moxel ReadFromMemory(IntPtr pSheetDoc)
        {
            CSheetDoc SheetDoc = new CSheetDoc(pSheetDoc);
            return ReadFromCSheetDoc(SheetDoc);
        }


        public class CSheetDoc : CObject
        {
            public CSheet Sheet = null;
            public PageSettings PageSettings = null;

            public int Length
            {
                get
                {
                    return 
                        Sheet.m_nCols * Sheet.m_nRows * (Marshal.SizeOf<CSheetFormat>() + 30); //Колонки * строки * размер форматной ячейки + Колонки * строки * 30 символов текста
                }
            }

            MoxelNative.SerializeDelegate _Serialize = MoxelNative.GetSerializer("?Serialize@CSheetDoc@@UAEXAAVCArchive@@@Z"); //CSheetDoc::Serialize(CSheetDoc *this, struct CArchive *Archive)

            public bool SaveToFile(string FileName)
            {
                using (FileStream fs = File.OpenWrite(FileName))
                {
                    using (CFile f = CFile.FromFileStream(fs))
                    {
                        try
                        {
                            using (CArchive Arch = new CArchive(f, this))
                            {
                                Serialize(Arch);
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            return false;
                        }
                    }
                }
            }

            public CSheetDoc(IntPtr pMem) : base(pMem)
            {
                Sheet = new CSheet(pMem + 0xB0);
                PageSettings = PageSettings.FromIntPtr(pMem + 0x354);
                Converter.PageSettings = PageSettings;
                //Debug.WriteLine($"CProfile7: {(pMem + 0x354).ToInt32():X8}");
            }

            public void Serialize(CArchive Arch)
            {
               _Serialize(this, Arch);
                Arch.Flush(); // Допишем оставшийся в буфере кусок данных.
            }
        }

        public class CSheet : CObject
        {
            MoxelNative.SerializeDelegate _Serialize = MoxelNative.GetSerializer("?Serialize@CSheet@@UAEXAAVCArchive@@@Z"); //CSheet::Serialize(CSheet *this, CArchive *a2)

            public int m_nCols = 0;                       //164h
            public int m_nRows = 0;                   //168h

            public CSheet(IntPtr pMem) : base(pMem)
            {
                //SheetDoc = GetMember<CSheetDoc>(0x21c);
                m_nCols = Marshal.ReadInt32(pMem, 0x164);
                m_nRows = Marshal.ReadInt32(pMem, 0x168);
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
           
            public CSheetDoc SheetDoc = null;
            public CSheetDoc SheetTemplate = null;

            public CTableOutputContext(IntPtr pMem) : base(pMem)
            {
                SheetDoc = GetMember<CSheetDoc>(0x20);
                SheetTemplate = GetMember<CSheetDoc>(0x2C);
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
                return obj.Pointer;
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

        public class CFile : CObject, IDisposable
        {
            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            public delegate IntPtr CFileConstructor(IntPtr pMem, IntPtr hFile);

            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            public delegate IntPtr CFileDestructor(IntPtr pMem);

            [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
            public delegate void CFile__Write(IntPtr _this, IntPtr lpBUf, int nCount);

            private static CFileConstructor _CFile = MFCNative.GetDelegate<CFileConstructor>(352); //??0CFile@@QAE@H@Z CFile::CFile(int HFile)

            private static CFileDestructor _CFileDestructor = MFCNative.GetDelegate<CFileDestructor>(665);

            private static CFile__Write OriginalWrite;

            public byte[] buffer = null;

            private Stream stream;

            int position = 0;

            public byte[] GetBufer()
            {
                unpatch();
                return buffer;
            }

            public Stream GetStream()
            {
                unpatch();
                stream.Seek(0, SeekOrigin.Begin);
                return stream;
            }

            public void StreamWrite(IntPtr _this, IntPtr lpBUf, int nCount)
            {
                if (_this == pObject)
                {
                    unsafe
                    {
                        byte* ptr = (byte*)lpBUf.ToPointer();
                        for (int i = 0; i < nCount; i++)
                        {
                            stream.WriteByte(ptr[i]);
                        }
                        position = (int)stream.Position;
                    }
                }
                else
                {
                    OriginalWrite(_this, lpBUf, nCount); //На случай записи другого CFile из другого потока.
                }
            }

            public void BufferWrite(IntPtr _this, IntPtr lpBUf, int nCount)
            {
                if (_this == pObject)
                {
                    if (buffer.Length < position + nCount + 1)
                        Array.Resize<byte>(ref buffer, position + nCount + 1);

                    unsafe
                    {
                        byte* ptr = (byte*)lpBUf.ToPointer();
                        for (int i = 0; i < nCount; i++)
                        {
                            stream.WriteByte(ptr[i]);
                        }
                    }
                    Marshal.Copy(lpBUf, buffer, position, nCount);
                    position += nCount;
                }
                else
                {
                    OriginalWrite(_this, lpBUf, nCount); //На случай записи другого CFile из другого потока.
                }
            }

            public CFile__Write WriteDelegate { get; set; }
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
                
                if(stream == null)
                    Handle = new CFile__Write(BufferWrite);
                else
                    Handle = new CFile__Write(StreamWrite);

                hh = GCHandle.Alloc(Handle);

                pWriteDelegate = Marshal.GetFunctionPointerForDelegate<CFile__Write>(Handle);

                uint OldProtection;

                WinApi.VirtualProtectEx(Process.GetCurrentProcess().Handle, FuncAddr, new IntPtr(4), 0x40, out OldProtection);
                Marshal.WriteIntPtr(FuncAddr, pWriteDelegate);
                WinApi.VirtualProtectEx(Process.GetCurrentProcess().Handle, FuncAddr, new IntPtr(4), OldProtection, out OldProtection);

                patched = true;
            }

            static CFile__Write Handle;
            private bool disposedValue;

            public void unpatch()
            {
                if(buffer != null)
                    Array.Resize<byte>(ref buffer, position);

                if (!patched)
                    return;

                uint OldProtection;
                WinApi.VirtualProtectEx(Process.GetCurrentProcess().Handle, FuncAddr, new IntPtr(4), 0x40, out OldProtection);
                Marshal.WriteIntPtr(FuncAddr, old_Func); // Вернем оригинальную функцию.
                WinApi.VirtualProtectEx(Process.GetCurrentProcess().Handle, FuncAddr, new IntPtr(4), OldProtection, out _);
                if(hh.IsAllocated)
                    hh.Free();
                patched = false;
            }

            private void InitPointers()
            {
                old_Func = Marshal.ReadIntPtr(pVTable, 0x40); // CFile::Write находитс по смещение 0x40 от начала таблицы виртуальных функций.
                OriginalWrite = Marshal.GetDelegateForFunctionPointer<CFile__Write>(old_Func);
                FuncAddr = new IntPtr((Int64)pVTable + 0x40);
                // перехватим CFile::Write()
                patch();
            }

            public CFile(IntPtr pMem, int bufferLength) : this(pMem)
            {
                buffer = new byte[bufferLength + 1];
                InitPointers();
            }

            public CFile(IntPtr pMem, Stream _stream) : this(pMem)
            {
                stream = _stream;
                InitPointers();
            }

            public CFile(IntPtr pMem) : base(pMem)
            {
                pVTable = Marshal.ReadIntPtr(pMem, 0);
            }

            public static CFile Create(int bufferLength)
            {
                IntPtr pMem = Marshal.AllocHGlobal(0x10);
                return new CFile(_CFile(pMem, IntPtr.Zero), bufferLength);
            }

            public static CFile FromStream(Stream stream)
            {
                IntPtr pMem = Marshal.AllocHGlobal(0x10);
                return new CFile(_CFile(pMem, IntPtr.Zero), stream);
            }

            public static CFile FromFileStream(FileStream stream)
            {
                IntPtr pMem = Marshal.AllocHGlobal(0x10);
                return new CFile(_CFile(pMem, stream.SafeFileHandle.DangerousGetHandle()));
            }

            ~CFile()
            {
                Dispose(false);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!disposedValue)
                {
                    if (disposing)
                    {
                        unpatch();
                        buffer = null;
                    }

                    try
                    {
                        _CFileDestructor(this);
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(pObject);
                        disposedValue = true;
                    }
                }
            }

            public void Dispose()
            {
                // Не изменяйте этот код. Разместите код очистки в методе "Dispose(bool disposing)".
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        public class CArchive: IDisposable
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

            CFile pfile;

            public IntPtr pDocument
            {
                set
                {
                    Marshal.WriteIntPtr(pObject, value); // CDocument по смещению 0 от начала объекта в памяти
                }
            }

            public IntPtr pObject;
            private bool disposedValue;

            [Flags]
            enum Mode : uint { store = 0, load = 1, bNoFlushOnDelete = 2, bNoByteSwap = 4 };


            public CArchive(CFile pfile, CSheetDoc CDocument)
            {
                int Length = CDocument.Sheet.m_nCols * CDocument.Sheet.m_nRows * 32;
                pObject = Marshal.AllocHGlobal(0x44);
                pObject = _CArchive(pObject, pfile, Mode.store, Length, IntPtr.Zero);
                this.pfile = pfile;
                pDocument = CDocument;
            }

            public void Flush()
            {
                flush(this);
                pfile.unpatch();
            }

            public static implicit operator IntPtr(CArchive obj)
            {
                return obj.pObject;
            }

            ~CArchive()
            {
                Dispose(false);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!disposedValue)
                {
                    if (disposing)
                    {
                        // TODO: освободить управляемое состояние (управляемые объекты)
                    }

                    try
                    {
                        _CArchiveDestructor(this);
                    }
                    catch
                    {
                    }
                    Marshal.FreeHGlobal(this);
                    disposedValue = true;
                }
            }

            public void Dispose()
            {
                // Не изменяйте этот код. Разместите код очистки в методе "Dispose(bool disposing)".
                Dispose(disposing: true);
                GC.SuppressFinalize(this);
            }
        }
    }

}
