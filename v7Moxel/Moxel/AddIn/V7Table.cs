// ============================================================================
// V7Table.cs — интеграция "SetTable" с заготовками Moxel (MemoryReader/MFCNative/
// MoxelNative/WinApi). Работает поверх MemoryReader.CObject/CTableOutputContext/
// CSheetDoc, но добавляет защиту памяти (IsBadReadPtr) — без неё чтение мусорного
// указателя в CObject-конструкторе даёт AccessViolation и НЕЛОВИМОЕ падение
// процесса 1С (CorruptedStateException в .NET Framework).
//
// ВЕРИФИЦИРОВАННАЯ ЦЕПОЧКА (по заголовкам 1С 7.7):
//   IUnknown* == IDispatch* == CBLExportContext*   (bkend.h:788, одиночное
//     наследование IDispatch, без вложенных интерфейсов)
//     +0x08  m_pCont ──► CBLContext* == CTableOutputContext (MOXEL.H:1425)
//       +0x20  m_pSheetDoc1 : CSheetDoc*
//       +0x24  m_nID        : UINT      → CTemplate7::GetDocument(m_nID) (Frame.h:1007)
//       +0x2C  m_pSheetDoc2 : CSheetDoc* (макет)
//     CSheetDoc (MOXEL.H:752):  +0xB0 CSheet m_Sheet   (+0x21C back-ptr)
//                               +0x354 CProfile7 m_Profile (PageSettings)
// ============================================================================

using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Moxel
{
    public class V7Table : IDisposable
    {
        // ------------------------------------------------------------------
        //  Верифицированные смещения
        // ------------------------------------------------------------------

        // CBLExportContext (bkend.h:788)
        const int OFF_EXPORTCTX_M_PCONT = 0x08;

        // CTableOutputContext (MOXEL.H:1425)
        const int OFF_TABCTX_SHEETDOC1 = 0x20;
        const int OFF_TABCTX_NID       = 0x24;
        const int OFF_TABCTX_SHEETDOC2 = 0x2C;

        // CSheetDoc (MOXEL.H:752)
        const int OFF_SHEETDOC_SHEET   = 0x0B0;
        const int OFF_SHEETDOC_PROFILE = 0x354;

        // CSheet (MOXEL.H:508) — относительно CSheet
        const int OFF_SHEET_BACKPTR    = 0x21C;
        const int OFF_SHEET_NCOLS      = 0x164;
        const int OFF_SHEET_NROWS      = 0x168;

        // CProfile7 (Frame.h:721) / CProfileEntry7 (Frame.h:709)
        const int OFF_PROFILE_ENTRYS   = 0x44;   // m_Entrys (CProfileEntryArr)
        const int OFF_ARR_PDATA        = 0x04;   // CArray::m_pData
        const int OFF_ARR_NSIZE        = 0x08;   // CArray::m_nSize (0x0C — это m_nMaxSize!)
        const int OFF_ENTRY_SIZE       = 0x14;
        const int OFF_ENTRY_TYPE       = 0x00;
        const int OFF_ENTRY_DATA1      = 0x04;
        const int OFF_ENTRY_NAME       = 0x0C;   // CString EntryName (== char*)

        // ------------------------------------------------------------------
        //  Состояние
        // ------------------------------------------------------------------

        /// <summary>RCW COM-объекта Таблицы. ХРАНИМ — пока жив он, живы
        /// CBLExportContext и таблица 1С, а с ним и сырые указатели.</summary>
        public object ComObject;

        /// <summary>CBLExportContext* (без владельческой ссылки).</summary>
        public IntPtr pExportContext = IntPtr.Zero;

        /// <summary>CTableOutputContext — BL-контекст Таблицы (с RTTI-проверкой).</summary>
        public MemoryReader.CTableOutputContext Context;

        /// <summary>CSheetDoc печатной формы (главный результат SetTable).</summary>
        public MemoryReader.CSheetDoc SheetDoc;

        /// <summary>CSheetDoc макета (m_pSheetDoc2) — если понадобится.</summary>
        public MemoryReader.CSheetDoc Maket;

        /// <summary>m_nID документа в реестре CTemplate7.</summary>
        public int DocumentID;

        public string LastError = string.Empty;

        public int Rows { get { return SheetDoc != null ? SheetDoc.Sheet.m_nRows : 0; } }
        public int Cols { get { return SheetDoc != null ? SheetDoc.Sheet.m_nCols : 0; } }

        MemSheet _mem;

        /// <summary>
        /// Прямое чтение ячеек/шрифтов/рисунков/секций из памяти
        /// (без сериализации в mxl). Пример: table.Mem.GetCell(2, 5).Text
        /// или string[,] grid = table.Mem.ReadTextGrid(0, 0);
        /// </summary>
        public MemSheet Mem
        {
            get
            {
                if (_mem == null && SheetDoc != null)
                    _mem = new MemSheet(this);
                return _mem;
            }
        }

        // ==================================================================
        //  SetTable — аналог CV7TableDocManager::SetTable(CValue**)
        // ==================================================================

        /// <summary>
        /// Принимает COM-объект Таблицы (1С передает IDispatch обертки
        /// CBLExportContext вокруг CTableOutputContext) и строит управляемые
        /// обертки над нативными объектами.
        /// </summary>
        public bool SetTable(object comTable)
        {
            Release();
            LastError = string.Empty;

            if (comTable == null) { LastError = "Аргумент равен null"; return false; }

            IntPtr pUnknown = IntPtr.Zero;
            try
            {
                // --- 1. IUnknown переданного объекта ------------------------
                // CBLExportContext реализует IDispatch напрямую (без вложенных
                // интерфейсов), поэтому IUnknown* == IDispatch* == адрес объекта.
                pUnknown = Marshal.GetIUnknownForObject(comTable);
                if (pUnknown == IntPtr.Zero) { LastError = "GetIUnknownForObject == 0"; return false; }

                if (!IsReadable(pUnknown, OFF_EXPORTCTX_M_PCONT + 4))
                { LastError = string.Format("Объект 0x{0:X8} не похож на CBLExportContext (память недоступна)", pUnknown.ToInt32()); return false; }

                // --- 2. m_pCont (+0x08) → CBLContext* -----------------------
                IntPtr pCont = Marshal.ReadIntPtr(pUnknown, OFF_EXPORTCTX_M_PCONT);

                // Защитный запас: у нестандартной обертки контекст может быть
                // в соседнем слоте (0x04 / 0x0C / 0x10).
                if (pCont == IntPtr.Zero || !LooksLikeBLContext(pCont))
                {
                    foreach (int off in new[] { 0x04, 0x0C, 0x10 })
                    {
                        if (!IsReadable(pUnknown, off + 4)) continue;
                        IntPtr cand = Marshal.ReadIntPtr(pUnknown, off);
                        if (cand != IntPtr.Zero && LooksLikeBLContext(cand)) { pCont = cand; break; }
                    }
                }

                // --- 3. Обертка контекста с RTTI-проверкой ------------------
                // CObject-конструктор сравнит имя CRuntimeClass с
                // "CTableOutputContext" и бросит исключение при несовпадении.
                Context = ConstructSafe<MemoryReader.CTableOutputContext>(pCont, "CTableOutputContext");
                if (Context == null) { LastError = LastErrorMessage; return false; }

                // --- 4. CSheetDoc — приоритет как у trad --------------------
                // GetSheetDoc(): CTemplate7::GetDocument(m_nID); fallback m_pSheetDoc1.
                DocumentID = Marshal.ReadInt32(pCont, OFF_TABCTX_NID);

                IntPtr pDoc = Template7GetDocument((uint)DocumentID);
                SheetDoc = ConstructSafe<MemoryReader.CSheetDoc>(pDoc, "CSheetDoc");

                if (SheetDoc == null)
                {
                    pDoc = Marshal.ReadIntPtr(pCont, OFF_TABCTX_SHEETDOC1);
                    SheetDoc = ConstructSafe<MemoryReader.CSheetDoc>(pDoc, "CSheetDoc");
                }

                if (SheetDoc == null) { LastError = LastErrorMessage; return false; }

                // Макет (опционально)
                Maket = ConstructSafe<MemoryReader.CSheetDoc>(
                    Marshal.ReadIntPtr(pCont, OFF_TABCTX_SHEETDOC2), "CSheetDoc");

                // --- 5. Валидация back-ptr CSheet::m_pSheetDoc (+0x21C) -----
                IntPtr pSheet = SheetDoc.Pointer + OFF_SHEETDOC_SHEET;
                if (IsReadable(pSheet, OFF_SHEET_BACKPTR + 4))
                {
                    IntPtr back = Marshal.ReadIntPtr(pSheet, OFF_SHEET_BACKPTR);
                    if (back != SheetDoc.Pointer)
                        Debug.WriteLine(string.Format(
                            "[V7Table] Внимание: CSheet::m_pSheetDoc 0x{0:X8} != CSheetDoc 0x{1:X8}",
                            back.ToInt32(), SheetDoc.Pointer.ToInt32()));
                }

                // Держим RCW — таблица 1С не умрет, пока жив этот объект.
                ComObject = comTable;
                pExportContext = pUnknown;

                Debug.WriteLine(string.Format(
                    "[V7Table] OK: nID={0} CSheetDoc=0x{1:X8} rows={2} cols={3} maket={4}",
                    DocumentID, SheetDoc.Pointer.ToInt32(), Rows, Cols,
                    Maket != null ? Maket.Pointer.ToInt32().ToString("X8") : "-"));
                return true;
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                Release();
                return false;
            }
            finally
            {
                // Ссылку на обертку отпускаем (не владельческая),
                // объект живет благодаря ComObject (RCW).
                if (pUnknown != IntPtr.Zero)
                    Marshal.Release(pUnknown);
            }
        }

        // ==================================================================
        //  Защищенное конструирование CObject-оберток
        //  (эквивалент CObject.FromComObject<T>, но с проверкой памяти
        //   ПЕРЕД входом в конструктор — иначе AV уронит процесс 1С)
        // ==================================================================

        string LastErrorMessage = string.Empty;

        T ConstructSafe<T>(IntPtr pMem, string expectedClassName) where T : MemoryReader.CObject
        {
            LastErrorMessage = string.Empty;
            if (pMem == IntPtr.Zero)
            {
                LastErrorMessage = string.Format("{0}: указатель NULL", expectedClassName);
                return null;
            }
            if (!IsReadable(pMem, 0x30))
            {
                LastErrorMessage = string.Format("{0}: память 0x{1:X8} недоступна", expectedClassName, pMem.ToInt32());
                return null;
            }

            // Предварительная проверка имени класса по RTTI (без конструирования)
            string actual = TryGetRuntimeClassName(pMem);
            if (actual != expectedClassName)
            {
                LastErrorMessage = string.Format("Ожидался {0}, получен '{1}' (0x{2:X8})",
                    expectedClassName, string.IsNullOrEmpty(actual) ? "<не MFC-объект>" : actual, pMem.ToInt32());
                return null;
            }

            try { return (T)Activator.CreateInstance(typeof(T), pMem); }
            catch (Exception ex)
            {
                LastErrorMessage = string.Format("{0}: {1}", expectedClassName, ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Имя MFC-класса по vtable[0] (GetRuntimeClass). Для MFC42 vftable
        /// начинается с GetRuntimeClass — см. BASIC.H (??_7CGetDoc7@@6B@).
        /// Не кидает исключений и не читает память без проверки.
        /// </summary>
        public static string TryGetRuntimeClassName(IntPtr pObj)
        {
            try
            {
                if (!IsReadable(pObj, 4)) return null;
                IntPtr vtable = Marshal.ReadIntPtr(pObj, 0);
                if (!IsReadable(vtable, 4)) return null;
                IntPtr pGetClass = Marshal.ReadIntPtr(vtable, 0);
                if (!IsReadable(pGetClass, 6)) return null;

                // MSVC: mov eax, imm32; ret  →  B8 xx xx xx xx C3
                if (Marshal.ReadByte(pGetClass, 0) != 0xB8) return null;
                if (Marshal.ReadByte(pGetClass, 5) != 0xC3) return null;

                IntPtr pRuntimeClass = Marshal.ReadIntPtr(pGetClass, 1);
                if (!IsReadable(pRuntimeClass, 4)) return null;

                IntPtr pName = Marshal.ReadIntPtr(pRuntimeClass, 0);
                if (!IsReadable(pName, 1)) return null;

                return Marshal.PtrToStringAnsi(pName);
            }
            catch { return null; }
        }

        static bool LooksLikeBLContext(IntPtr pCont)
        {
            // Классы BL-контекстов — MFC CObject-наследники с RTTI.
            string name = TryGetRuntimeClassName(pCont);
            return !string.IsNullOrEmpty(name);
        }

        // ==================================================================
        //  CTemplate7::GetDocument (Frame.h:1007)
        //  static CDocument* GetDocument(unsigned int);
        //  → ?GetDocument@CTemplate7@@SAPAVCDocument@@I@Z
        // ==================================================================

        static IntPtr _pGetDocument = IntPtr.Zero;

        public static IntPtr Template7GetDocument(uint nID)
        {
            if (_pGetDocument == IntPtr.Zero && !ResolveGetDocument())
                return IntPtr.Zero;

            try
            {
                return _GetDocument((uint)nID);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[V7Table] GetDocument(" + nID + "): " + ex.Message);
                return IntPtr.Zero;
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]   // static __cdecl
        delegate IntPtr _GetDocumentDelegate(uint nID);
        static _GetDocumentDelegate _GetDocument;

        const string GETDOCUMENT_ENTRY = "?GetDocument@CTemplate7@@SAPAVCDocument@@I@Z";

        static bool ResolveGetDocument()
        {
            foreach (string dll in new[] { "bkend.dll", "blang.dll", "moxel.dll" })
            {
                IntPtr h = WinApi.GetModuleHandle(dll);
                if (h == IntPtr.Zero) continue;

                IntPtr p = WinApi.GetProcAddress(h, GETDOCUMENT_ENTRY);
                if (p != IntPtr.Zero)
                {
                    _pGetDocument = p;
                    _GetDocument = Marshal.GetDelegateForFunctionPointer<_GetDocumentDelegate>(p);
                    Debug.WriteLine("[V7Table] " + GETDOCUMENT_ENTRY + " → " + dll);
                    return true;
                }
            }
            Debug.WriteLine("[V7Table] " + GETDOCUMENT_ENTRY + " не найден");
            return false;
        }

        // ==================================================================
        //  Работа с содержимым
        // ==================================================================

        /// <summary>Сериализовать таблицу в mxl-образ (нативный Serialize
        /// moxel.dll + перехват CFile::Write из ваших заготовок).</summary>
        public Moxel ReadMoxel()
        {
            if (SheetDoc == null)
                throw new InvalidOperationException("Таблица не установлена. Сначала SetTable().");
            return MemoryReader.ReadFromMemory(SheetDoc.Pointer);
        }

        /// <summary>
        /// Вызвать метод Таблицы 1С по имени через RCW — 1С сама смаршалит
        /// вызов в CTableOutputContext::CallAsProc/CallAsFunc.
        /// Пример: InvokeTableMethod("ПараметрыСтраницы", "A4", "Альбомная", 100, 20, 20, 20, 20);
        /// </summary>
        public object InvokeTableMethod(string name, params object[] args)
        {
            if (ComObject == null)
                throw new InvalidOperationException("Таблица не установлена. Сначала SetTable().");

            return ComObject.GetType().InvokeMember(
                name, BindingFlags.InvokeMethod, null, ComObject, args ?? new object[0]);
        }

        /// <summary>Таблица.ПараметрыСтраницы(...) — родной механизм 1С.
        /// Параметры страницы в 7.7 хранятся записями CProfile7
        /// (CSheetDoc+0x354), а не бинарной структурой — меняйте их этим
        /// методом, а не записью в чужую память.</summary>
        public void SetPageSetup(string paper, string orientation, int scale,
                                 int left, int right, int top, int bottom)
        {
            InvokeTableMethod("ПараметрыСтраницы", paper, orientation, scale,
                              left, right, top, bottom);
        }

        /// <summary>Таблица.Напечатать(1).</summary>
        public void Print(int copies) { InvokeTableMethod("Напечатать", copies); }

        /// <summary>Таблица.Показать().</summary>
        public void Show(string title) { InvokeTableMethod("Показать", title ?? string.Empty); }

        // ==================================================================
        //  Диагностика: CProfile7 @ CSheetDoc+0x354
        //  Параметры страницы лежат здесь записями CProfileEntry7.
        //  Дамп покажет фактические ключи на вашей сборке 1С.
        // ==================================================================

        public void DumpProfile()
        {
            if (SheetDoc == null) return;

            IntPtr pProfile = SheetDoc.Pointer + OFF_SHEETDOC_PROFILE;
            if (!IsReadable(pProfile, OFF_PROFILE_ENTRYS + 0x10)) return;

            IntPtr pEntrys = pProfile + OFF_PROFILE_ENTRYS;
            IntPtr pData = Marshal.ReadIntPtr(pEntrys, OFF_ARR_PDATA);
            int nSize = Marshal.ReadInt32(pEntrys, OFF_ARR_NSIZE);
            if (pData == IntPtr.Zero || nSize <= 0) return;

            Debug.WriteLine(string.Format("=== CProfile7 @ 0x{0:X8}: {1} записей ===",
                pProfile.ToInt32(), nSize));

            for (int i = 0; i < nSize && i < 256; i++)
            {
                IntPtr pEntry = Marshal.ReadIntPtr(pData, i * 4);
                if (pEntry == IntPtr.Zero || !IsReadable(pEntry, OFF_ENTRY_SIZE)) continue;

                // CString (MFC42 ANSI) — указатель на символьные данные
                IntPtr pChars = Marshal.ReadIntPtr(pEntry, OFF_ENTRY_NAME);
                string name = (pChars != IntPtr.Zero && IsReadable(pChars, 1))
                            ? (Marshal.PtrToStringAnsi(pChars) ?? string.Empty)
                            : string.Empty;

                Debug.WriteLine(string.Format("  [{0,3}] type={1} data1={2} (0x{2:X8}) name='{3}'",
                    i, Marshal.ReadInt32(pEntry, OFF_ENTRY_TYPE),
                    Marshal.ReadInt32(pEntry, OFF_ENTRY_DATA1), name));
            }
        }

        // ==================================================================
        //  Утилиты / время жизни
        // ==================================================================

        static bool IsReadable(IntPtr p, int size)
        {
            if (p == IntPtr.Zero || size <= 0) return false;
            uint addr = (uint)p.ToInt32();
            if (addr < 0x00010000 || addr >= 0x7FFF0000) return false;
            return IsBadReadPtr(p, (uint)size) == 0;
        }

        [DllImport("kernel32.dll")]
        static extern int IsBadReadPtr(IntPtr lp, uint ucb);

        /// <summary>
        /// Освободить таблицу. Сырые указатели обнуляем; RCW отпускаем —
        /// время жизни нативных объектов 1С снова регулирует 1С.
        /// ВАЖНО: не обращайтесь к SheetDoc.Context после того, как в 1С
        /// таблица уничтожена (Сформировать/закрытие формы).
        /// </summary>
        public void Release()
        {
            Context = null;
            SheetDoc = null;
            _mem = null;   // старый CSheet умрет вместе с таблицей 1С
            Maket = null;
            pExportContext = IntPtr.Zero;
            DocumentID = 0;
            ComObject = null;
        }

        public void Dispose() { Release(); }
        ~V7Table() { Release(); }
    }
}
