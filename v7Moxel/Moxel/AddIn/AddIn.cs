using System;
using System.Runtime.InteropServices;
using AddIn;
using System.Reflection;
using System.Collections;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using static Moxel.MemoryReader;
using System.IO.MemoryMappedFiles;
using System.Diagnostics;

namespace Moxel
{
    public sealed class AliasAttribute : Attribute
    {
        public string RussianName { get; set; }
        public AliasAttribute(string alias)
        {
            RussianName = alias;
        }
    }

    public static class WinApi
    {
        [DllImport("Kernel32", EntryPoint = "GetModuleHandleW", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string dllname);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr GetProcAddress(IntPtr HModule, [MarshalAs(UnmanagedType.LPStr), In] string funcName);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr GetProcAddress(IntPtr HModule, int ordinal);

        [DllImport("kernel32.dll")]
        public static extern bool VirtualProtectEx(IntPtr hProcess, IntPtr lpAddress,  IntPtr dwSize, uint flNewProtect, out uint lpflOldProtect);
    }

    public static class MoxelNative
    {

       [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate void SerializeDelegate(IntPtr pObj, IntPtr pArch);

        static IntPtr hMoxel = WinApi.GetModuleHandle("moxel.dll");

        public static SerializeDelegate GetSerializer(string EntryPoint)
        {
            IntPtr ProcAddress = WinApi.GetProcAddress(hMoxel, EntryPoint);
             return Marshal.GetDelegateForFunctionPointer<SerializeDelegate>(ProcAddress);
        }
    }

    public static class MFCNative
    {

        static IntPtr hMFC = WinApi.GetModuleHandle("MFC42.dll");

        [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate IntPtr pGetRuntimeClass(IntPtr pObj);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate IntPtr _CreateObject(IntPtr pMem);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate IntPtr _GetBaseClass();

        public static T GetDelegate<T>(string EntryPoint)
        {
            return Marshal.GetDelegateForFunctionPointer<T>(WinApi.GetProcAddress(hMFC, EntryPoint));
        }

        public static T GetDelegate<T>(int Ordinal)
        {
            return Marshal.GetDelegateForFunctionPointer<T>(WinApi.GetProcAddress(hMFC, Ordinal));
        }
    }


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
                catch(Exception ex)
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
                byte tt = Marshal.ReadByte(pGetClass, 0);

                if (tt == 0xB8) // MOV EAX
                {
                    MFCNative.pGetRuntimeClass GetRuntimeClass = Marshal.GetDelegateForFunctionPointer<MFCNative.pGetRuntimeClass>(pGetClass);
                    RuntimeClass = Marshal.PtrToStructure<CRuntimeClass>(GetRuntimeClass(pObject));
                    if (this.GetType().Name != RuntimeClass.ClassName)
                        throw new Exception( $"Указатель на объект отличного от {this.GetType().Name} типа");
                }
                else
                     throw new Exception($"Указатель не на CRuntimeClass ({this.GetType().Name} pGetClass:BYTE [{pGetClass.ToInt32():X8}] = {tt:X8})");
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

    [ComVisible(true)]
    [InterfaceType( ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1EAE378F-C315-4B49-980C-A9A40792E78C")]
    internal interface IConverter
    {
        [Alias("Присоединить")]
        void Attach(object Table);

        [Alias("ПрисоединитьИзПамяти")]
        void AttachToMemory(object Table);

        [Alias("Записать")]
        string Save(string filename, SaveFormat format);
    }


    [ComVisible(true)]
    [Guid("2DF0622D-BC0A-4C30-8B7D-ACB66FB837B6")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Description("Конвертер MOXEL")]
    [ProgId("AddIn.Moxel.Converter")]
    public class Converter : IInitDone, ILanguageExtender, IConverter
    {

        /// <summary>ProgID COM-объекта компоненты</summary>
        string AddInName = "Moxel.Converter";
        //string AddInName = "Таблица";

        /// <summary>Указатель на IDispatch</summary>
        protected object connect1c;

        /// <summary>Вызов событий 1С</summary>
        protected IAsyncEvent asyncEvent;

        /// <summary>Статусная строка 1С</summary>
        protected IStatusLine statusLine;

        /// <summary>Сообщения об ошибках 1С</summary>
        protected IErrorLog errorLog;

        private Type[] allInterfaceTypes;  // Коллекция интерфейсов
        private MethodInfo[] allMethodInfo;  // Коллекция методов
        private PropertyInfo[] allPropertyInfo; // Коллекция свойств

        private Hashtable nameToNumber; // метод - идентификатор
        private Hashtable numberToName; // идентификатор - метод
        private Hashtable numberToParams; // количество параметров метода
        private Hashtable numberToRetVal; // имеет возвращаемое значение (является функцией)
        private Hashtable propertyNameToNumber; // свойство - идентификатор
        private Hashtable propertyNumberToName; // идентификатор - свойство
        private Hashtable numberToMethodInfoIdx; // номер метода - индекс в массиве методов
        private Hashtable propertyNumberToPropertyInfoIdx; // номер свойства - индекс в массиве свойств

        Moxel mxl;

        protected void PostException(Exception ex)
        {
            if (errorLog == null)
                return;

            var info = new System.Runtime.InteropServices.ComTypes.EXCEPINFO
            {
                wCode = 1006,
                bstrDescription = $"{AddInName}: ошибка {ex.GetType()} : {ex.Message}", 
                bstrSource = AddInName, 
                scode = 1
            };

            errorLog.AddError("", ref info);
        }



        public void AttachToMemory(object Table)
        {
            var TableObject = CObject.FromComObject<CTableOutputContext>(Table);
            var Sheet = TableObject.Sheet;

            //while (Marshal.ReleaseComObject(Table) > 0) { }
            //Marshal.FinalReleaseComObject(Table);

            CFile f = CFile.FromHFile(IntPtr.Zero);
            CArchive Arch = new CArchive(f, Sheet.SheetDoc);
            Sheet.SheetDoc.Serialize(Arch); 
            Arch.Flush();
            Arch = null;
            byte[] buffer = f.GetBufer();
            f = null;
            mxl = new Moxel(ref buffer);
        }



        public void Attach(object Table)
        {
            string tempfile = Path.GetTempFileName();
            File.Delete(tempfile);
            tempfile += ".mxl";
            object[] param = { tempfile, "mxl" };
            var tt = Table.GetType().InvokeMember("Write", BindingFlags.InvokeMethod, null, Table, param);

            //while (Marshal.ReleaseComObject(Table) > 0){}
            //Marshal.FinalReleaseComObject(Table);

            if (File.Exists(tempfile))
                mxl = new Moxel(tempfile);

            File.Delete(tempfile);
        }

        public string Save(string filename, SaveFormat format)
        {
            mxl.SaveAs(filename, format);
            return filename;
        }


        #region IInitDone
        HRESULT IInitDone.Init([MarshalAs(UnmanagedType.IDispatch)] object connection)
        {
            connect1c = connection;
            statusLine = (IStatusLine)connection;
            asyncEvent = (IAsyncEvent)connection;
            errorLog = (IErrorLog)connection;
            return HRESULT.S_OK;
        }

        HRESULT IInitDone.GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
        {
            info[0] = 2000;
            return HRESULT.S_OK;
        }

        HRESULT IInitDone.Done()
        {
            
            if (connect1c != null)
            {
                while (Marshal.ReleaseComObject(asyncEvent) > 0) { };
                Marshal.FinalReleaseComObject(asyncEvent);
                asyncEvent = null;

                while (Marshal.ReleaseComObject(statusLine) > 0) { };
                Marshal.FinalReleaseComObject(statusLine);
                statusLine = null;

                while (Marshal.ReleaseComObject(errorLog) > 0) { };
                Marshal.FinalReleaseComObject(errorLog);
                statusLine = null;

                while (Marshal.ReleaseComObject(connect1c) > 0) { };
                Marshal.FinalReleaseComObject(connect1c);
                connect1c = null;
            }
            Marshal.CleanupUnusedObjectsInCurrentContext();

            mxl = null;

            GC.Collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return HRESULT.S_OK;
        }
        #endregion


#region ILAnguageExtender
        HRESULT ILanguageExtender.CallAsFunc(int methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                retValue = allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
            }
            catch (Exception e)
            {
                PostException(e.InnerException);
            }
            return HRESULT.S_OK;
        }

        HRESULT ILanguageExtender.CallAsProc(int methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref dynamic[] pParams)
        {
            try
            {
                allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
                
            }
            catch (Exception e)
            {
                PostException(e.InnerException);
            }
            return HRESULT.S_OK;
        }


        int ILanguageExtender.FindMethod([MarshalAs(UnmanagedType.BStr)] string methodName)
        {
            if (nameToNumber.ContainsKey(methodName.ToUpper()))
                return (int)nameToNumber[methodName.ToUpper()];
            else
            return -1;
            
        }

        HRESULT ILanguageExtender.FindProp([MarshalAs(UnmanagedType.BStr)] string propName, ref Int32 propNum)
        {
            if (propertyNameToNumber.ContainsKey(propName.ToUpper()))
            {
                propNum = (Int32)propertyNameToNumber[propName.ToUpper()];
                return HRESULT.S_OK;
            }

            propNum = -1;
            return HRESULT.S_FALSE;
        }



        HRESULT ILanguageExtender.GetMethodName(int methodNum, int methodAlias, [MarshalAs(UnmanagedType.BStr)] ref string methodName)
        {
            if (numberToName.ContainsKey(methodNum))
            {
                methodName = (String)numberToName[methodNum];
                return HRESULT.S_OK;
            }
            return HRESULT.S_FALSE;

        }

        HRESULT ILanguageExtender.GetNMethods(ref Int32 pMethods)
        {
            pMethods = allMethodInfo.Length;
            return HRESULT.S_OK;
        }

        HRESULT ILanguageExtender.GetNParams(int methodNum, ref int pParams)
        {
            if (numberToParams.ContainsKey(methodNum))
            {
                pParams = (Int32)numberToParams[methodNum];
                return HRESULT.S_OK;
            }

            pParams = -1;
            return HRESULT.S_FALSE;
        }

        public HRESULT GetNProps(ref int props)
        {
            props = (Int32)propertyNameToNumber.Count;
            return HRESULT.S_OK;
        }

        public HRESULT GetParamDefValue(int methodNum, int paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue)
        {
            return HRESULT.S_OK;
        }

        public HRESULT GetPropName(int propNum, int propAlias, [MarshalAs(UnmanagedType.BStr)] ref string propName)
        {
            if (propertyNumberToName.ContainsKey(propNum))
            {
                propName = (String)propertyNumberToName[propNum];
                return HRESULT.S_OK;
            }
            return HRESULT.S_FALSE;
        }

        public HRESULT GetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            propVal = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].GetValue(this, null);
            return HRESULT.S_OK;
        }

        HRESULT ILanguageExtender.HasRetVal(int methodNum, ref bool retValue)
        {
            if (numberToRetVal.ContainsKey(methodNum))
            {
                retValue = (bool)numberToRetVal[methodNum];
                return HRESULT.S_OK;
            }

            return HRESULT.S_FALSE;
        }


        HRESULT ILanguageExtender.IsPropReadable(int propNum, ref bool propRead)
        {
            propRead = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanRead;
            return HRESULT.S_OK;
        }

        public HRESULT IsPropWritable(int propNum, ref bool propWrite)
        {
            propWrite = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanWrite;
            return HRESULT.S_OK;
        }

        HRESULT ILanguageExtender.RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName)
        {
            try
            {
                Type type = this.GetType();

                allInterfaceTypes = type.GetInterfaces();
                allMethodInfo = type.GetMethods();
                allPropertyInfo = type.GetProperties();

                // Хэш-таблицы с именами методов и свойств компоненты
                nameToNumber = new Hashtable();
                numberToName = new Hashtable();
                numberToParams = new Hashtable();
                numberToRetVal = new Hashtable();
                propertyNameToNumber = new Hashtable();
                propertyNumberToName = new Hashtable();
                numberToMethodInfoIdx = new Hashtable();
                propertyNumberToPropertyInfoIdx = new Hashtable();

                int Identifier = 0;

                foreach (Type interfaceType in allInterfaceTypes)
                {
                    // Интересуют только методы в пользовательских интерфейсах, стандартные пропускаем
                    if ( interfaceType.Name.Equals("IDisposable")
                      || interfaceType.Name.Equals("IManagedObject")
                      || interfaceType.Name.Equals("IRemoteDispatch")
                      || interfaceType.Name.Equals("IServicedComponentInfo")
                      || interfaceType.Name.Equals("IInitDone")
                      || interfaceType.Name.Equals("ILanguageExtender"))
                    {
                        continue;
                    };

                    // Обработка методов интерфейса
                    MethodInfo[] interfaceMethods = interfaceType.GetMethods();
                    foreach (MethodInfo interfaceMethodInfo in interfaceMethods)
                    {
                        string alias = ((AliasAttribute)Attribute.GetCustomAttribute(interfaceMethodInfo, typeof(AliasAttribute))).RussianName.ToUpper();

                        nameToNumber.Add(interfaceMethodInfo.Name.ToUpper(), Identifier);
                        numberToName.Add(Identifier, interfaceMethodInfo.Name);
                        numberToParams.Add(Identifier, interfaceMethodInfo.GetParameters().Length);
                        numberToRetVal.Add(Identifier, (interfaceMethodInfo.ReturnType != typeof(void)));
                        if (!string.IsNullOrWhiteSpace(alias))
                        {
                            nameToNumber.Add(alias, Identifier);
                        }
                        Identifier++;
                    }

                    // Обработка свойств интерфейса
                    PropertyInfo[] interfaceProperties = interfaceType.GetProperties();
                    foreach (PropertyInfo interfacePropertyInfo in interfaceProperties)
                    {
                        string alias = ((AliasAttribute)Attribute.GetCustomAttribute(interfacePropertyInfo, typeof(AliasAttribute))).RussianName.ToUpper();

                        propertyNameToNumber.Add(interfacePropertyInfo.Name, Identifier);

                        propertyNumberToName.Add(Identifier, interfacePropertyInfo.Name);

                        if (!string.IsNullOrWhiteSpace(alias))
                        {
                            propertyNameToNumber.Add(alias, Identifier);
                        }

                        Identifier++;
                    }
                }

                // Отображение номера метода на индекс в массиве
                foreach (DictionaryEntry entry in numberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < allMethodInfo.Length; ii++)
                    {
                        if (allMethodInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            numberToMethodInfoIdx.Add(entry.Key, ii);
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        throw new COMException("Метод " + entry.Value.ToString() + " не реализован");
                }

                // Отображение номера свойства на индекс в массиве
                foreach (DictionaryEntry entry in propertyNumberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < allPropertyInfo.Length; ii++)
                    {
                        if (allPropertyInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            propertyNumberToPropertyInfoIdx.Add(entry.Key, ii);
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        throw new COMException("Свойство " + entry.Value.ToString() + " не реализовано");
                }

                // Компонент инициализирован успешно. Возвращаем имя компонента.
                extensionName = AddInName;
            }
            catch (Exception e)
            {
                PostException(e);
            }

            return HRESULT.S_OK;
            
        }

        public HRESULT SetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].SetValue(this, propVal, null);
            return HRESULT.S_OK;
        }
        #endregion
    }
}
