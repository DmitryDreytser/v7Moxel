using AddIn;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using v7Moxel.Moxel.ExcelWriter;
using static Moxel.MemoryReader;



using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Moxel
{


    [ComVisible(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1EAE378F-C315-4B49-980C-A9A40792E78C")]
    internal interface IConverter
    {
        [Alias("Присоединить")]
        void Attach(object Table);

        [Alias("Загрузить")]
        void Load(string FileNAme);

        [Alias("ПерехватВключен")]
        int IsWrapped { get;}

        [Alias("ЗагрузитьИзПамяти")]
        void ReadFromMemory(object Table);

        [Alias("Записать")]
        string Save(string filename, SaveFormat format);

        [Alias("ПерехватитьЗапись")]
        int WrapSaveAs(int doWrap = 1);

        [Alias("ОписаниеОшибки")]
        string GetErrorDescription();

        [Alias("СтекОшибки")]
        string GetErrorStackTrace();

        [Alias("КоличествоПотоковКонвертации")]
        int MaxDegreeOfParallelism { get; set; }
    }


    [ComVisible(true)]    
    [Guid("2DF0622D-BC0A-4C30-8B7D-ACB66FB837B6")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(IConverter))]
    [Description("Конвертер MOXEL")]
    [ProgId( "AddIn.Moxel.Converter")]
    public class Converter : AddIn, IConverter
    {
        public static void RaiseExtRuntimeError(string ErrorMessage)
        {
                var bkendUi = GetBkendUi();

                var vfTable = Marshal.ReadIntPtr(bkendUi);

                var DoMessageLine = Marshal.GetDelegateForFunctionPointer<dDoMessageLine>(Marshal.ReadIntPtr(vfTable + 0xC));

                DoMessageLine.Invoke(bkendUi, Marshal.StringToCoTaskMemAnsi(ErrorMessage), MessageMarker.RedErr);
            
        }

        public static Moxel mxl;
        public static PageSettings PageSettings = null;
        public static CTableOutputContext TableObject = null;
        static int ObjectCount = 0;

        static Converter()
        {
            SaveWrapper.Wrap(true);
        }

        public Converter()
        {
            
        }

        public int IsWrapped { get => SaveWrapper.isWraped ? 1 : 0; }
        public int MaxDegreeOfParallelism { get => ExcelWriter.MaxDegreeOfParallelism ; set
            {
                ExcelWriter.MaxDegreeOfParallelism = value;
            }
        }
        public int WrapSaveAs(int doWrap = 1) => SaveWrapper.Wrap(doWrap == 1);
        public void ReadFromMemory(object Table)
        {
            try
            {
                TableObject = CObject.FromComObject<CTableOutputContext>(Table);
                PageSettings = TableObject.SheetDoc.PageSettings;

                while (Marshal.ReleaseComObject(Table) > 0) { }
                Marshal.FinalReleaseComObject(Table);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                mxl?.Dispose();
                mxl = null;
            }
        }

        public unsafe class V7TableDocManager
        {
            /// <summary>
            /// Ориентация страницы
            /// </summary>
            public enum PageOrientation : ushort
            {
                Portrait = 1,  // Книжная
                Landscape = 2  // Альбомная
            }

            /// <summary>
            /// Размер бумаги (стандартные значения Win32 / 1C)
            /// </summary>
            public enum PagePaperSize : ushort
            {
                Letter = 1,
                A3 = 8,
                A4 = 9,
                A5 = 11,
                // Другие значения совпадает с DMPAPER_* из Win32 API
            }

            /// <summary>
            /// Внутренняя структура параметров страницы Moxel (CSheetDoc / CPageSetup)
            /// Выравнивание по байтам соответствуют MSVC 6.0 (Pack = 4)
            /// </summary>

            [StructLayout(LayoutKind.Sequential, Pack = 4)]
            public struct MoxcelPageSetup
            {
                public ushort Orientation;    // 1 - Portrait, 2 - Landscape
                public ushort PaperSize;      // ID размера бумаги (DMPAPER_A4 = 9)
                public ushort Scale;          // Масштаб в процентах (например, 100)
                public ushort FitToPageWidth; // Умещать по ширине (в страницах)
                public ushort FitToPageHeight;// Умещать по высоте (в страницах)
                public ushort Flags;          // Флаги (FitToPage, PrintGridLines, etc.)

                // Поля в миллиметрах (или десятых долях мм в зависимости от версии)
                public int MarginLeft;        // Левое поле
                public int MarginTop;         // Верхнее поле
                public int MarginRight;       // Правое поле
                public int MarginBottom;      // Нижнее поле

                public int HeaderMargin;      // Отступ колонтитула сверху
                public int FooterMargin;      // Отступ колонтитула снизу
            }


            // Адрес документа таблицы в памяти (CSheetDoc*)
            private IntPtr _pSheetDoc = IntPtr.Zero;

            // Адрес самой таблицы в памяти (CSheet*)
            private IntPtr _pSheet = IntPtr.Zero;


            // Смещение CPageSetup внутри CSheetDoc (для 1С 7.7 release / 27)
            private const int PAGE_SETUP_OFFSET = 0x84;


            public static void LogContextStructure(IntPtr pDispatch)
            {
                Debug.WriteLine($"=== DUMP START ===");
                Debug.WriteLine($"pDispatch Native Address: 0x{pDispatch.ToInt32():X8}");

                // 1. Дампим сам pDispatch (первые 8 слотов / 32 байта)
                if (IsValidReadPtr(pDispatch, 32))
                {
                    IntPtr* pFields = (IntPtr*)pDispatch.ToPointer();
                    Debug.WriteLine("\n--- pDispatch Fields ---");
                    for (int i = 0; i < 8; i++)
                    {
                        IntPtr ptr = pFields[i];
                        string rtti = GetMfcClassName(ptr);
                        string rttiMfcOffset = GetMfcClassName(new IntPtr(ptr.ToInt32() - 8));

                        Debug.WriteLine($"pDispatch[+{i * 4:X2}] = 0x{ptr.ToInt32():X8} | RTTI direct: '{rtti}' | RTTI (-8): '{rttiMfcOffset}'");
                    }
                }

                // 2. Сканируем окрестности pDispatch и ищем CTableOutputContext
                for (int offset = -64; offset <= 64; offset += 4)
                {
                    IntPtr candidate = new IntPtr(pDispatch.ToInt32() + offset);
                    string className = GetMfcClassName(candidate);

                    if (className == "CTableOutputContext")
                    {
                        Debug.WriteLine($"\n[FOUND] CTableOutputContext at offset (pDispatch {offset:+#;-#;+0}) -> 0x{candidate.ToInt32():X8}");

                        // 3. Дампим внутренности CTableOutputContext (первые 40 байт) для поиска m_nID
                        Debug.WriteLine("--- CTableOutputContext Memory Dump ---");
                        int* pInts = (int*)candidate.ToPointer();
                        for (int j = 0; j < 10; j++)
                        {
                            Debug.WriteLine($"Context[+{j * 4:X2}] = 0x{pInts[j]:X8} ({pInts[j]})");
                        }
                    }
                }
                Debug.WriteLine($"=== DUMP END ===\n");
            }
            /// <summary>
            /// Аналог CV7TableDocManager::SetTable
            /// Принимает COM-объект Таблицы 1С 7.7 и извлекает native-указатели.
            /// </summary>
            /// <param name="comTable">Объект Таблица (полученный через OLE / ВК 1С)</param>
            /// <returns>true, если объект успешно получен и валиден</returns>
            public bool SetTable(object comTable)
            {
                _pSheetDoc = IntPtr.Zero;
                _pSheet = IntPtr.Zero;

                if (comTable == null)
                    return false;

                IntPtr pUnknown = IntPtr.Zero;
                IntPtr pDispatch = IntPtr.Zero;

                try
                {
                    // 1. Получаем IUnknown / IDispatch указатель из COM-объекта C#
                    pUnknown = Marshal.GetIUnknownForObject(comTable);
                    if (pUnknown == IntPtr.Zero)
                        return false;

                    // Запрашиваем IDispatch
                    Guid iidIDispatch = new Guid("00020400-0000-0000-C000-000000000046");
                    int hr = Marshal.QueryInterface(pUnknown, ref iidIDispatch, out pDispatch);

                    if (hr != 0 || pDispatch == IntPtr.Zero)
                    {
                        // Если IDispatch не поддерживается, пробуем работать прямо с pUnknown
                        pDispatch = pUnknown;
                    }

                    LogContextStructure(pDispatch);

                    // 2. Аналог CCmdTarget::FromIDispatch(pDispatch)
                    // В MFC 4.2 (32-bit 1С 7.7) смещение IDispatch относительно CCmdTarget равно 8 байтам.
                    const int MFC_INTERFACE_PART_OFFSET = 8;

                    IntPtr pCmdTarget = FindCmdTarget(pDispatch);

                    //string className = GetMfcClassName(pCmdTarget);

                    //// Ожидаемый класс таблицы 1С 7.7 — CSheetDoc
                    //if (className != "CSheetDoc")
                    //{
                    //    // Если передана не таблица (например, C1CBookDoc или другое окно/объект)
                    //    return false;
                    //}

                    //// 3. Проверка валидности полученного указателя
                    //if (!IsValidMemoryPtr(pCmdTarget))
                    //    return false;

                    static IntPtr FindCmdTarget(IntPtr pDispatch)
                    {
                        // MFC CCmdTarget vtable в 1С 7.7 обычно начинается с адреса в диапазоне DLL 1С (mfc42.dll / bkend.dll / moxcel.dll)
                        // Пробуем сканировать диапазон от pDispatch - 32 байт до pDispatch + 32 байт
                        for (int offset = -32; offset <= 32; offset += 4)
                        {
                            IntPtr candidate = new IntPtr(pDispatch.ToInt32() + offset);

                            if (GetMfcClassName(candidate) == "CTableOutputContext")
                            {
                                // Нашли точное смещение для данного типа объекта!
                                return candidate;
                            }
                        }
                        return IntPtr.Zero;
                    }

                    if (pCmdTarget == IntPtr.Zero)
                    {
                        // Если не нашли напрямую, пробуем развернуть IDispatch через разбор интерфейса
                        // В 1С 7.7 OLE-обертка таблицы хранит указатель на CSheetDoc по смещению +0x08 или +0x0C
                        unsafe
                        {
                            IntPtr* pTablePtr = (IntPtr*)pDispatch.ToPointer();
                            // Пробуем прочитать вложенный указатель
                            if (IsValidReadPtr((IntPtr)pTablePtr, 16))
                            {
                                for (int i = 0; i < 4; i++)
                                {
                                    IntPtr nestedPtr = pTablePtr[i];
                                    pCmdTarget = FindCmdTarget(nestedPtr);
                                    if (pCmdTarget != IntPtr.Zero) break;
                                }
                            }
                        }
                    }

                    _pSheetDoc = pCmdTarget;

                    // 4. Получение CSheet* из CSheetDoc*
                    // В 1С 7.7 объект CSheet лежит внутри CSheetDoc по смещению 0xB0 (176 байт)
                    // Или по альтернативному вычислению: ((DWORD*)pSheetDoc) + 0x2C (что равен 0x2C * 4 = 176 = 0xB0)
                    const int CSHEET_OFFSET = 0xB0;
                    _pSheet = new IntPtr(_pSheetDoc.ToInt32() + CSHEET_OFFSET);

                    return true;
                }
                catch
                {
                    _pSheetDoc = IntPtr.Zero;
                    _pSheet = IntPtr.Zero;
                    return false;
                }
                finally
                {
                    if (pDispatch != IntPtr.Zero && pDispatch != pUnknown)
                        Marshal.Release(pDispatch);
                    if (pUnknown != IntPtr.Zero)
                        Marshal.Release(pUnknown);
                }
            }

            /// <summary>
            /// Указатель на CSheetDoc
            /// </summary>
            public IntPtr CSheetDocPointer => _pSheetDoc;

            /// <summary>
            /// Указатель на CSheet
            /// </summary>
            public IntPtr CSheetPointer => _pSheet;


            /// <summary>
            /// Получить параметры страницы / печати из таблицы
            /// </summary>
            public MoxcelPageSetup GetPageSetup()
            {
                if (_pSheetDoc == IntPtr.Zero)
                    throw new InvalidOperationException("Таблица не инициализирована. Вызовите SetTable.");

                IntPtr pPageSetup = new IntPtr(_pSheetDoc.ToInt32() + PAGE_SETUP_OFFSET);

                // Читаем структуру прямо из памяти C++ процесса 1С
                return Marshal.PtrToStructure<MoxcelPageSetup>(pPageSetup);
            }

            /// <summary>
            /// Записать новые параметры страницы / печати в таблицу
            /// </summary>
            public void SetPageSetup(MoxcelPageSetup setup)
            {
                if (_pSheetDoc == IntPtr.Zero)
                    throw new InvalidOperationException("Таблица не инициализирована. Вызовите SetTable.");

                IntPtr pPageSetup = new IntPtr(_pSheetDoc.ToInt32() + PAGE_SETUP_OFFSET);

                // Записываем обновленную структуру обратно в память 1С
                Marshal.StructureToPtr(setup, pPageSetup, fDeleteOld: false);
            }

            /// <summary>
            /// Удобные свойства-обертки для быстрой работы
            /// </summary>
            public PageOrientation Orientation
            {
                get => (PageOrientation)GetPageSetup().Orientation;
                set
                {
                    var setup = GetPageSetup();
                    setup.Orientation = (ushort)value;
                    SetPageSetup(setup);
                }
            }

            /// <summary>
            /// Поля таблицы (в миллиметрах)
            /// </summary>
            public (int Left, int Top, int Right, int Bottom) Margins
            {
                get
                {
                    var s = GetPageSetup();
                    return (s.MarginLeft, s.MarginTop, s.MarginRight, s.MarginBottom);
                }
                set
                {
                    var setup = GetPageSetup();
                    setup.MarginLeft = value.Left;
                    setup.MarginTop = value.Top;
                    setup.MarginRight = value.Right;
                    setup.MarginBottom = value.Bottom;
                    SetPageSetup(setup);
                }
            }

            /// <summary>
            /// Масштаб печати (%)
            /// </summary>
            public ushort Scale
            {
                get => GetPageSetup().Scale;
                set
                {
                    var setup = GetPageSetup();
                    setup.Scale = value;
                    SetPageSetup(setup);
                }
            }

            /// <summary>
            /// Простейшая проверка валидности указателя памяти
            /// </summary>
            private static bool IsValidMemoryPtr(IntPtr ptr)
            {
                if (ptr == IntPtr.Zero)
                    return false;

                // В 32-битных Windows пользовательский адрес должен быть в пределах юзер-пейса
                uint addr = (uint)ptr.ToInt32();
                if (addr < 0x00010000 || addr >= 0x7FFF0000)
                    return false;

                // Дополнительная проверка: пытаемся прочитать vtable
                return Win32.IsBadReadPtr(ptr, (uint)IntPtr.Size) == 0;
            }

            private static bool IsValidReadPtr(IntPtr ptr, int size)
            {
                if (ptr == IntPtr.Zero) return false;
                uint addr = (uint)ptr.ToInt32();
                if (addr < 0x00010000 || addr >= 0x7FFF0000) return false;

                return Win32.IsBadReadPtr(ptr, (uint)size) == 0;
            }

            /// <summary>
            /// Читает имя MFC-класса объекта через его CRuntimeClass / VTable
            /// </summary>
            public static string GetMfcClassName(IntPtr pMfcObject)
            {
                if (!IsValidReadPtr(pMfcObject, 4)) return string.Empty;

                try
                {
                    // vtable находится по первому адресу объекта (*pMfcObject)
                    IntPtr vtable = *(IntPtr*)pMfcObject.ToPointer();
                    if (!IsValidReadPtr(vtable, 4)) return string.Empty;

                    // В MFC первыми слотами vtable идут:
                    // [0] GetRuntimeClass()
                    IntPtr pGetRuntimeClass = *(IntPtr*)vtable.ToPointer();
                    if (!IsValidReadPtr(pGetRuntimeClass, 4)) return string.Empty;

                    // Вызываем GetRuntimeClass() через stdcall (указатель 'this' передается первой инструкцией или ECX)
                    // В MSVC (MFC) метод GetRuntimeClass() обычно просто возвращает статический адрес CRuntimeClass:
                    // mov eax, offset CSheetDoc::classCSheetDoc; ret
                    // Попробуем прочитать CRuntimeClass напрямую, вызвав скомпилированную функцию или распарсив её return:

                    // Самый надежный и безопасный способ без выполнения чужого кода:
                    // В 32-битном MSVC GetRuntimeClass() состоит из:
                    // B8 [XX XX XX XX] (mov eax, pRuntimeClass) -> C3 (ret)
                    byte* code = (byte*)pGetRuntimeClass.ToPointer();
                    if (code[0] == 0xB8 && code[5] == 0xC3) // Opcode MOV EAX, imm32; RET
                    {
                        IntPtr pRuntimeClass = *(IntPtr*)(code + 1);
                        var @class =  ReadClassNameFromRuntimeClass(pRuntimeClass);
                        Debug.WriteLine($"got RTTI: {@class}");
                    }
                    return String.Empty;
                    // Запасной способ: прямое вызов функции GetRuntimeClass via Delegate
                    var getRuntimeClassDelegate = (GetRuntimeClassDelegate)Marshal.GetDelegateForFunctionPointer(
                        pGetRuntimeClass,
                        typeof(GetRuntimeClassDelegate)
                    );

                    IntPtr pClassStruct = getRuntimeClassDelegate(pMfcObject);
                    return ReadClassNameFromRuntimeClass(pClassStruct);
                }
                catch
                {
                    return string.Empty;
                }
            }

            /// <summary>
            /// Структура CRuntimeClass из MFC 4.2:
            /// LPCSTR m_lpszClassName; (смещение 0x00)
            /// int m_nObjectSize;     (смещение 0x04)
            /// UINT m_wSchema;         (смещение 0x08)
            /// ...
            /// </summary>
            private static string ReadClassNameFromRuntimeClass(IntPtr pRuntimeClass)
            {
                if (!IsValidReadPtr(pRuntimeClass, 4)) return string.Empty;

                // По смещению +00 лежит указатель на ANSI-строку m_lpszClassName
                IntPtr pClassNameStr = *(IntPtr*)pRuntimeClass.ToPointer();
                if (!IsValidReadPtr(pClassNameStr, 2)) return string.Empty;

                return Marshal.PtrToStringAnsi(pClassNameStr) ?? string.Empty;
            }

            // Делегат для вызова `virtual CRuntimeClass* GetRuntimeClass() const`
            [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
            private delegate IntPtr GetRuntimeClassDelegate(IntPtr pThis);

            // COM IDispatch GUID определение
            [ComImport, Guid("00020400-0000-0000-C000-000000000464"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            private interface IDispatch { }

            private static class Win32
            {
                [DllImport("kernel32.dll")]
                public static extern int IsBadReadPtr(IntPtr lp, uint ucb);
            }
        }

        public void Attach(object Table)
        {
            mxl?.Dispose();
            try
            {

                var manager = new V7TableDocManager();
                manager.SetTable(Table);

                var pSheetDoc = manager.CSheetDocPointer;

                string tempfile = Path.GetTempFileName();
                if (pSheetDoc != IntPtr.Zero)
                {
                    var doc = new CSheetDoc(pSheetDoc);


                }
                File.Delete(tempfile);
                tempfile += ".mxl";
                object[] param = { tempfile, "mxl" };
                var tt = Table.GetType().InvokeMember("Write", BindingFlags.InvokeMethod, null, Table, param);

                if (File.Exists(tempfile))
                    mxl = new Moxel(tempfile);

                File.Delete(tempfile);

                while (Marshal.ReleaseComObject(Table) > 0) { }
                Marshal.FinalReleaseComObject(Table);

            }
            catch (Exception ex)
            {
                while (Marshal.ReleaseComObject(Table) > 0) { }
                Marshal.FinalReleaseComObject(Table);
                throw ex.InnerException;
            }
        }

        public void Load(string FileNAme)
        {

            if (!File.Exists(FileNAme))
                throw new Exception($"Файл {FileNAme} не найден.");

            mxl?.Dispose();
            mxl = new Moxel(FileNAme);
        }

        public string Save(string filename, SaveFormat format)
        {
            try
            {
                if (mxl == null)
                {
                    if (TableObject != null)
                        if (TableObject.SheetDoc.Length < 1024 * 1024 * 2 || !SaveWrapper.CanSaveExternal)
                            mxl = ReadFromCSheetDoc(TableObject.SheetDoc);
                        else
                        {
                            string tmpFileName = Path.GetTempFileName();
                            TableObject.SheetDoc.SaveToFile(tmpFileName);
                            if (SaveWrapper.SaveExternal(tmpFileName, filename).Result == 1)
                            {
                                File.Delete(tmpFileName);
                                return filename;
                            }
                            else
                                throw new Exception("Ошибка записи.");
                        }
                    else
                        throw new Exception("Таблица не загружена.");

                }

                if (mxl != null)
                {
                    mxl.SaveAs(filename, format);
                    return filename;
                }
                else
                {
                    throw new Exception("Таблица не загружена.");
                }
            }
            finally
            {
                mxl?.Dispose();
                mxl = null;
                GC.Collect();
                GC.Collect();
            }
        }


        #region AddIn events

        public string GetErrorDescription()
        {
            return this.ErrorDescription;
        }

        public string GetErrorStackTrace()
        {
            return this.ErrorStackTrace;
        }

        protected override void OnInit()
        {
            ObjectCount++;
        }

        protected override HRESULT OnRegister()
        {
            //ObjectCount++;
            return HRESULT.S_OK;
        }

        protected override void OnDone()
        {
            if (--ObjectCount == 0) // При выгрузке последнего объекта из памяти отключим перехват, если он есть.
                SaveWrapper.Wrap(false);

            mxl = null;
        }



        #endregion

    }
}
