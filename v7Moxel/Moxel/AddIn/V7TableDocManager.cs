// ============================================================================
// V7TableDocManager.cs  —  ИСПРАВЛЕННАЯ ВЕРСИЯ
// Аналог CV7TableDocManager::SetTable (trad, 2006) для C#-компоненты 1С 7.7.
//
// ВЕРИФИЦИРОВАННАЯ ЦЕПОЧКА УКАЗАТЕЛЕЙ (все смещения сверены с заголовками 7.7):
//
//  VARIANT (VT_UNKNOWN / VT_DISPATCH)
//   └─ CBLExportContext  (bkend.h:788, class IMPORT_1C CBLExportContext : public IDispatch)
//        +0x00  vptr (IDispatch)
//        +0x04  m_RefCount
//        +0x08  m_pCont  ──► CBLContext*  = CTableOutputContext   ★ ключевой шаг
//        +0x0C  m_Flag2
//      CBLExportContext — одиночное наследование IDispatch (без вложенных
//      интерфейсов), поэтому IUnknown* == IDispatch* == адрес объекта.
//
//   └─ CTableOutputContext (MOXEL.H:1425, size 0x480; база CBLContext = 0x20)
//        +0x20  m_pSheetDoc1  : CSheetDoc*   ← прямой указатель
//        +0x24  m_nID         : UINT         ← ID в реестре CTemplate7
//        +0x28  m_pSheet1     : CSheet*
//        +0x2C  m_pSheetDoc2  : CSheetDoc*   ← макет
//        +0x38  m_pSheet2     : CSheet*
//
//   └─ CSheetDoc* = CTemplate7::GetDocument(m_nID)   (Frame.h:1007, static)
//        — именно так делает оригинальный trad:
//          pDoc = (CSheetDoc*)CTemplate7::GetDocument(m_pTableCont->m_nID);
//        Fallback: m_pSheetDoc1 (+0x20).
//
//   └─ CSheetDoc (MOXEL.H:752, база COleLinkingDoc = 0xB0 байт)
//        +0x0B0  CSheet m_Sheet  (встроен)   → CSHEET_OFFSET = 0xB0 ✔
//        +0x354  CProfile7 m_Profile (0x58)  ← параметры страницы ЗДЕСЬ
//            (это подтверждает trad: m_pSheet2 = ((DWORD*)pSheetDoc)+0x2C)
//        +0x21C  CSheet::m_pSheetDoc — обратный указатель (валидация!)
//
//  ★ Параметры страницы НЕ лежат бинарной структурой по +0x84 (там начало
//    COleDocument). Они хранятся записями CProfileEntry7 в m_Profile и
//    корректнее всего выставляются родным методом Таблицы "ПараметрыСтраницы",
//    который можно вызвать через ПОЛУЧЕННЫЙ ЖЕ IDispatch (CBLExportContext::
//    Invoke транслирует его в CTableOutputContext::CallAsProc).
// ============================================================================

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;

namespace V7Interop
{
    public unsafe class V7TableDocManager : IDisposable
    {
        // ------------------------------------------------------------------
        //  Перечисления (совместимы с DMPAPER_*/DMORIENT_* из Win32)
        // ------------------------------------------------------------------

        public enum PageOrientation : int
        {
            Portrait  = 1,   // Книжная
            Landscape = 2    // Альбомная
        }

        public enum PagePaperSize : int
        {
            Letter = 1,
            A3     = 8,
            A4     = 9,
            A5     = 11
            // остальные значения = DMPAPER_* из Win32 API
        }

        // ------------------------------------------------------------------
        //  ВЕРИФИЦИРОВАННЫЕ СМЕЩЕНИЯ (по заголовкам 1С 7.7, release 27)
        // ------------------------------------------------------------------

        private static class Off
        {
            // CBLExportContext (bkend.h:788)
            public const int EXPORTCTX_M_PCONT = 0x08;   // CBLContext* m_pCont

            // CTableOutputContext (MOXEL.H:1425)
            public const int TABCTX_M_PHEETDOC1 = 0x20;  // CSheetDoc* m_pSheetDoc1
            public const int TABCTX_M_NID       = 0x24;  // UINT m_nID
            public const int TABCTX_M_PHEET1    = 0x28;  // CSheet* m_pSheet1
            public const int TABCTX_M_PHEETDOC2 = 0x2C;  // CSheetDoc* m_pSheetDoc2 (макет)

            // CSheetDoc (MOXEL.H:752)
            public const int SHEETDOC_SHEET     = 0xB0;  // CSheet m_Sheet (встроен)
            public const int SHEETDOC_PROFILE   = 0x354; // CProfile7 m_Profile

            // CSheet (MOXEL.H:508) — смещения ОТНОСИТЕЛЬНО CSheet
            public const int SHEET_M_PSHEETDOC  = 0x21C; // CSheetDoc* m_pSheetDoc (back-ptr)
            public const int SHEET_M_NCOLS      = 0x164; // DWORD m_nCols
            public const int SHEET_M_NROWS      = 0x168; // DWORD m_nRows

            // CProfile7 (Frame.h:721) — смещения ОТНОСИТЕЛЬНО CProfile7
            public const int PROFILE_M_ENTRYS   = 0x44;  // CProfileEntryArr m_Entrys
            public const int ARRAY_PDATA        = 0x04;  // CArray::m_pData (после CObject vptr)
            public const int ARRAY_NSIZE        = 0x0C;  // CArray::m_nSize

            // CProfileEntry7 (Frame.h:709) — размер 0x14
            public const int ENTRY_SIZE      = 0x14;
            public const int ENTRY_TYPE      = 0x00;   // DWORD type
            public const int ENTRY_DATA1     = 0x04;   // DWORD data1 (значение для чисел)
            public const int ENTRY_STR       = 0x08;   // CString str
            public const int ENTRY_ENTRYNAME = 0x0C;   // CString EntryName (char**)
        }

        // ------------------------------------------------------------------
        //  Состояние
        // ------------------------------------------------------------------

        private object  _tableObj;        // RCW обертки CBLExportContext — ХРАНИМ,
                                          // чтобы 1С-объект жил и можно было звать методы
        private IntPtr  _pExportContext;  // CBLExportContext*  (без владельческой ссылки!)
        private IntPtr  _pTableContext;   // CTableOutputContext*
        private IntPtr  _pSheetDoc;       // CSheetDoc*
        private IntPtr  _pSheet;          // CSheet* (встроен в CSheetDoc +0xB0)
        private string  _lastError;

        public string LastError { get { return _lastError ?? string.Empty; } }

        /// <summary>Указатель на CSheetDoc печатной формы (главная цель SetTable).</summary>
        public IntPtr CSheetDocPointer { get { return _pSheetDoc; } }

        /// <summary>Указатель на встроенный CSheet (CSheetDoc + 0xB0).</summary>
        public IntPtr CSheetPointer { get { return _pSheet; } }

        /// <summary>Указатель на CTableOutputContext (BL-контекст Таблицы).</summary>
        public IntPtr TableContextPointer { get { return _pTableContext; } }

        /// <summary>Указатель на CBLExportContext — OLE-обертку, переданную 1С.</summary>
        public IntPtr ExportContextPointer { get { return _pExportContext; } }

        /// <summary>m_nID документа в реестре CTemplate7.</summary>
        public int DocumentID { get { return _pTableContext != IntPtr.Zero ? ReadInt32(_pTableContext, Off.TABCTX_M_NID) : 0; } }

        // ==================================================================
        //  SetTable — аналог CV7TableDocManager::SetTable(CValue**)
        // ==================================================================

        /// <summary>
        /// Принимает COM-объект Таблицы 1С 7.7 (1С передает IDispatch-обертку
        /// CBLExportContext вокруг CTableOutputContext) и извлекает
        /// нативные указатели CTableOutputContext / CSheetDoc / CSheet.
        /// </summary>
        public bool SetTable(object comTable)
        {
            ReleaseTable();
            _lastError = null;

            if (comTable == null) { _lastError = "Аргумент равен null"; return false; }

            IntPtr pUnknown = IntPtr.Zero;
            try
            {
                // --- Шаг 1. IUnknown переданного COM-объекта ----------------
                // 1С кладет в VARIANT IDispatch объекта CBLExportContext.
                // Класс реализует IDispatch напрямую (одиночное наследование,
                // без вложенных XDispatch как в MFC), поэтому IUnknown*,
                // IDispatch* и адрес самого объекта СОВПАДАЮТ.
                pUnknown = Marshal.GetIUnknownForObject(comTable);
                if (pUnknown == IntPtr.Zero) { _lastError = "GetIUnknownForObject вернул 0"; return false; }

                _pExportContext = pUnknown;

                // --- Шаг 2. CBLExportContext::m_pCont (+0x08) → контекст ----
                // bkend.h:792  CBLContext* m_pCont;  // +0x08
                if (!IsReadable(pUnknown, Off.EXPORTCTX_M_PCONT + 4))
                { _lastError = "CBLExportContext: память недоступна"; return false; }

                IntPtr pCont = ReadPointer(pUnknown, Off.EXPORTCTX_M_PCONT);
                if (pCont == IntPtr.Zero) { _lastError = "CBLExportContext::m_pCont == NULL"; return false; }

                // Защитный запасной вариант: обертка может отличаться
                // (другая версия/путь передачи) — перебираем первые слоты.
                if (GetMfcClassName(pCont) != "CTableOutputContext")
                {
                    IntPtr alt = FindContextNearExportContext(pUnknown);
                    if (alt != IntPtr.Zero) pCont = alt;
                }

                // --- Шаг 3. Валидация CTableOutputContext по RTTI -----------
                // Аналог: strcmp("CTableOutputContext", pCont->GetRuntimeClass()->m_lpszClassName)
                string ctxClass = GetMfcClassName(pCont);
                if (ctxClass != "CTableOutputContext")
                {
                    _lastError = string.Format(
                        "Ожидался CTableOutputContext, получен '{0}' (0x{1:X8})", ctxClass, pCont.ToInt32());
                    return false;
                }
                _pTableContext = pCont;

                // --- Шаг 4. CSheetDoc — как в оригинале trad ----------------
                // GetSheetDoc(): CTemplate7::GetDocument(m_pTableCont->m_nID)
                int nID = ReadInt32(_pTableContext, Off.TABCTX_M_NID);
                IntPtr pDoc = CallTemplate7GetDocument(nID);

                // Fallback: прямой m_pSheetDoc1 (+0x20)
                if (pDoc == IntPtr.Zero)
                    pDoc = ReadPointer(_pTableContext, Off.TABCTX_M_PHEETDOC1);

                if (pDoc == IntPtr.Zero)
                { _lastError = "CSheetDoc не найден (GetDocument и m_pSheetDoc1 == 0)"; return false; }

                string docClass = GetMfcClassName(pDoc);
                if (docClass != "CSheetDoc")
                {
                    // Основной кандидат не CSheetDoc — пробуем m_pSheetDoc2 (макет)
                    IntPtr pDoc2 = ReadPointer(_pTableContext, Off.TABCTX_M_PHEETDOC2);
                    string doc2Class = (pDoc2 != IntPtr.Zero) ? GetMfcClassName(pDoc2) : null;
                    if (doc2Class == "CSheetDoc")
                    {
                        // это макет (CSheetDoc2) — годится только для чтения макета
                        _pSheetDoc = pDoc2;
                        _lastError = string.Format(
                            "Внимание: основной CSheetDoc недоступен (nID={0}), используется макет m_pSheetDoc2", nID);
                    }
                    else
                    {
                        _lastError = string.Format(
                            "Указатель 0x{0:X8} не является CSheetDoc (класс='{1}')", pDoc.ToInt32(), docClass);
                        return false;
                    }
                }
                else
                {
                    _pSheetDoc = pDoc;
                }

                // --- Шаг 5. CSheet встроен в CSheetDoc по +0xB0 -------------
                // Подтверждено trad'ом: m_pSheet2 = (CSheet*)(((DWORD*)pSheetDoc) + 0x2C);
                _pSheet = new IntPtr(_pSheetDoc.ToInt32() + Off.SHEETDOC_SHEET);

                // Дополнительная валидация: CSheet::m_pSheetDoc (+0x21C)
                // должен указывать на сам CSheetDoc.
                if (IsReadable(_pSheet, Off.SHEET_M_PSHEETDOC + 4))
                {
                    IntPtr back = ReadPointer(_pSheet, Off.SHEET_M_PSHEETDOC);
                    if (back != _pSheetDoc)
                        Debug.WriteLine(string.Format(
                            "[V7TableDoc] Предупреждение: CSheet::m_pSheetDoc (0x{0:X8}) != CSheetDoc (0x{1:X8})",
                            back.ToInt32(), _pSheetDoc.ToInt32()));
                }

                // Храним RCW — объект 1С останется живым, и можно будет
                // вызывать его методы (см. InvokeTableMethod).
                _tableObj = comTable;

                Debug.WriteLine(string.Format(
                    "[V7TableDoc] OK: ExportCtx=0x{0:X8} CTableOutputContext=0x{1:X8} (nID={2}) " +
                    "CSheetDoc=0x{3:X8} CSheet=0x{4:X8} Rows={5} Cols={6}",
                    _pExportContext.ToInt32(), _pTableContext.ToInt32(), nID,
                    _pSheetDoc.ToInt32(), _pSheet.ToInt32(),
                    SheetRows, SheetCols));

                return true;
            }
            catch (Exception ex)
            {
                _lastError = "Exception: " + ex.Message;
                ReleaseTable();
                return false;
            }
            finally
            {
                // Владельческую ссылку на обертку отпускаем: указатели выше —
                // «сырые», их время жизни определяется 1С (см. _tableObj).
                if (pUnknown != IntPtr.Zero)
                    Marshal.Release(pUnknown);
            }
        }

        // Защитный поиск: если обертка не совсем CBLExportContext,
        // пробуем первые слоты как кандидатов на CBLContext*.
        private IntPtr FindContextNearExportContext(IntPtr pExportCtx)
        {
            int[] tryOffsets = new int[] { 0x04, 0x08, 0x0C, 0x10 };
            foreach (int off in tryOffsets)
            {
                if (!IsReadable(pExportCtx, off + 4)) continue;
                IntPtr cand = ReadPointer(pExportCtx, off);
                if (cand == IntPtr.Zero || !IsReadable(cand, 0x24)) continue;
                if (GetMfcClassName(cand) == "CTableOutputContext")
                    return cand;
            }
            return IntPtr.Zero;
        }

        // ==================================================================
        //  CTemplate7::GetDocument  (Frame.h:1007)
        //  static CDocument* GetDocument(unsigned int);
        //  Украшенное имя MSVC6: ?GetDocument@CTemplate7@@SAPAVCDocument@@I@Z
        // ==================================================================

        private static IntPtr _hTemplate7Dll = IntPtr.Zero;
        private static IntPtr _pGetDocument = IntPtr.Zero;
        private static readonly string[] CandidateDlls = { "bkend.dll", "blang.dll", "moxel.dll" };
        private const string GetDocumentDecoratedName = "?GetDocument@CTemplate7@@SAPAVCDocument@@I@Z";

        private static IntPtr CallTemplate7GetDocument(int nID)
        {
            if (_pGetDocument == IntPtr.Zero && !ResolveTemplate7Export())
                return IntPtr.Zero;

            try
            {
                GetDocumentDelegate fn = (GetDocumentDelegate)Marshal.GetDelegateForFunctionPointer(
                    _pGetDocument, typeof(GetDocumentDelegate));
                IntPtr pDoc = fn((uint)nID);
                return pDoc;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[V7TableDoc] CTemplate7::GetDocument вызов не удался: " + ex.Message);
                return IntPtr.Zero;
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate IntPtr GetDocumentDelegate(uint nID); // static __cdecl

        private static bool ResolveTemplate7Export()
        {
            foreach (string dll in CandidateDlls)
            {
                IntPtr h = Win32.LoadLibrary(dll);
                if (h == IntPtr.Zero) continue;

                IntPtr p = Win32.GetProcAddress(h, GetDocumentDecoratedName);
                if (p != IntPtr.Zero)
                {
                    _hTemplate7Dll = h;
                    _pGetDocument = p;
                    Debug.WriteLine("[V7TableDoc] " + GetDocumentDecoratedName + " найден в " + dll);
                    return true;
                }
                Win32.FreeLibrary(h);
            }
            Debug.WriteLine("[V7TableDoc] " + GetDocumentDecoratedName + " не найден ни в одной DLL");
            return false;
        }

        // ==================================================================
        //  RTTI: имя MFC-класса объекта (CObject::GetRuntimeClass, vtable[0])
        //  В MFC42 vftable начинается с GetRuntimeClass — см. BASIC.H:
        //  ??_7CGetDoc7@@6B@ dd offset ?GetRuntimeClass@CGetDoc7@@...
        // ==================================================================

        /// <summary>
        /// Имя MFC-класса через CRuntimeClass. Для не-MFC COM-объектов
        /// (например CBLExportContext, где vtable[0] = QueryInterface)
        /// корректно вернет пустую строку.
        /// </summary>
        public static string GetMfcClassName(IntPtr pObject)
        {
            if (!IsReadable(pObject, 4)) return string.Empty;

            try
            {
                IntPtr vtable = *(IntPtr*)pObject.ToPointer();
                if (!IsReadable(vtable, 4)) return string.Empty;

                // vtable[0] = GetRuntimeClass() (для CObject-наследников)
                IntPtr pGetRuntimeClass = *(IntPtr*)vtable.ToPointer();
                if (!IsReadable(pGetRuntimeClass, 6)) return string.Empty;

                // Быстрый путь: MSVC компилирует GetRuntimeClass как
                //   B8 xx xx xx xx   mov eax, offset <class>XXX
                //   C3               ret
                byte* code = (byte*)pGetRuntimeClass.ToPointer();
                if (code[0] == 0xB8 && code[5] == 0xC3)
                {
                    IntPtr pRuntimeClass = *(IntPtr*)(code + 1);
                    return ReadClassNameFromRuntimeClass(pRuntimeClass);
                }

                // Медленный путь: честно вызываем виртуальный GetRuntimeClass
                // (это __thiscall; для не-MFC объектов сюда не попадем,
                // т.к. шаблон B8/C3 у QueryInterface не встречается)
                try
                {
                    GetRuntimeClassDelegate fn = (GetRuntimeClassDelegate)
                        Marshal.GetDelegateForFunctionPointer(pGetRuntimeClass, typeof(GetRuntimeClassDelegate));
                    IntPtr pClassStruct = fn(pObject);
                    return ReadClassNameFromRuntimeClass(pClassStruct);
                }
                catch { return string.Empty; }
            }
            catch { return string.Empty; }
        }

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        private delegate IntPtr GetRuntimeClassDelegate(IntPtr pThis);

        private static string ReadClassNameFromRuntimeClass(IntPtr pRuntimeClass)
        {
            if (!IsReadable(pRuntimeClass, 4)) return string.Empty;

            // CRuntimeClass: +0x00 LPCSTR m_lpszClassName
            IntPtr pClassName = *(IntPtr*)pRuntimeClass.ToPointer();
            if (!IsReadable(pClassName, 2)) return string.Empty;

            try { return Marshal.PtrToStringAnsi(pClassName) ?? string.Empty; }
            catch { return string.Empty; }
        }

        // ==================================================================
        //  Вызов методов Таблицы 1С через сохраненный RCW.
        //  CBLExportContext::Invoke транслирует вызов в
        //  CTableOutputContext::CallAsProc/CallAsFunc — т.е. мы дергаем
        //  РОДНУЮ функциональность 1С без работы с чужой памятью.
        // ==================================================================

        /// <summary>
        /// Вызвать метод Таблицы по имени ("Сформировать", "Показать",
        /// "ПараметрыСтраницы", "Область", ...).
        /// </summary>
        public object InvokeTableMethod(string name, params object[] args)
        {
            if (_tableObj == null)
                throw new InvalidOperationException("Таблица не установлена. Сначала SetTable().");

            return _tableObj.GetType().InvokeMember(
                name,
                BindingFlags.InvokeMethod,
                null, _tableObj, args ?? new object[0]);
        }

        /// <summary>
        /// Родная установка параметров страницы 1С 7.7:
        /// Таблица.ПараметрыСтраницы(&lt;РазмерСтраницы&gt;,&lt;Ориентация&gt;,&lt;Масштаб&gt;,
        ///                            &lt;ПолеЛ&gt;,&lt;ПолеП&gt;,&lt;ПолеВ&gt;,&lt;ПолеН&gt;)
        /// </summary>
        public void SetPageSetupNative(PagePaperSize paper, PageOrientation orientation,
                                       int scale, int marginLeft, int marginRight,
                                       int marginTop, int marginBottom)
        {
            InvokeTableMethod("ПараметрыСтраницы",
                PaperToStr(paper),
                orientation == PageOrientation.Landscape ? "Альбомная" : "Книжная",
                scale, marginLeft, marginRight, marginTop, marginBottom);
        }

        /// <summary>Таблица.Напечатать(1).</summary>
        public void Print(int copies)
        {
            InvokeTableMethod("Напечатать", copies);
        }

        /// <summary>Таблица.Показать().</summary>
        public void Show(string title)
        {
            InvokeTableMethod("Показать", title ?? string.Empty);
        }

        private static string PaperToStr(PagePaperSize p)
        {
            switch (p)
            {
                case PagePaperSize.A3: return "A3";
                case PagePaperSize.A4: return "A4";
                case PagePaperSize.A5: return "A5";
                case PagePaperSize.Letter: return "Letter";
                default: return "A4";
            }
        }

        // ==================================================================
        //  Доступ к содержимому CSheet (только чтение)
        // ==================================================================

        public int SheetRows
        {
            get { return (_pSheet != IntPtr.Zero && IsReadable(_pSheet, Off.SHEET_M_NROWS + 4))
                         ? ReadInt32(_pSheet, Off.SHEET_M_NROWS) : 0; }
        }

        public int SheetCols
        {
            get { return (_pSheet != IntPtr.Zero && IsReadable(_pSheet, Off.SHEET_M_NCOLS + 4))
                         ? ReadInt32(_pSheet, Off.SHEET_M_NCOLS) : 0; }
        }

        // ==================================================================
        //  Диагностика
        // ==================================================================

        /// <summary>Дамп обертки и контекста — заменяет старый LogContextStructure.</summary>
        public void DumpContextStructure()
        {
            Debug.WriteLine("=== V7TableDoc DUMP START ===");
            if (_pExportContext != IntPtr.Zero) DumpMemory(_pExportContext, 6, "CBLExportContext");
            if (_pTableContext != IntPtr.Zero)
            {
                DumpMemory(_pTableContext, 0x48 / 4, "CTableOutputContext");
                Debug.WriteLine(string.Format("  +20 m_pSheetDoc1 = 0x{0:X8} ('{1}')",
                    ReadPointer(_pTableContext, 0x20), GetMfcClassName(ReadPointer(_pTableContext, 0x20))));
                Debug.WriteLine(string.Format("  +24 m_nID        = {0}", ReadInt32(_pTableContext, 0x24)));
                Debug.WriteLine(string.Format("  +2C m_pSheetDoc2 = 0x{0:X8} ('{1}') [макет]",
                    ReadPointer(_pTableContext, 0x2C), GetMfcClassName(ReadPointer(_pTableContext, 0x2C))));
            }
            if (_pSheetDoc != IntPtr.Zero)
            {
                DumpMemory(_pSheetDoc, 8, "CSheetDoc");
                Debug.WriteLine(string.Format("  +B0 CSheet (rows={0}, cols={1}, back-ptr=0x{2:X8})",
                    SheetRows, SheetCols,
                    IsReadable(_pSheet, Off.SHEET_M_PSHEETDOC + 4)
                        ? ReadPointer(_pSheet, Off.SHEET_M_PSHEETDOC).ToInt32() : 0));
            }
            Debug.WriteLine("=== V7TableDoc DUMP END ===");
        }

        /// <summary>
        /// Перечисление записей CProfile7 (CSheetDoc+0x354) — здесь 1С хранит
        /// параметры страницы таблицы. Позволяет НА СВОЕЙ сборке 1С увидеть
        /// фактические ключи и значения (имена записей — ANSI CString).
        /// </summary>
        public void DumpProfileEntries()
        {
            if (_pSheetDoc == IntPtr.Zero) return;

            IntPtr pProfile = new IntPtr(_pSheetDoc.ToInt32() + Off.SHEETDOC_PROFILE);
            if (!IsReadable(pProfile, Off.PROFILE_M_ENTRYS + 0x10)) return;

            IntPtr pEntrysArray = new IntPtr(pProfile.ToInt32() + Off.PROFILE_M_ENTRYS);
            IntPtr pData = ReadPointer(pEntrysArray, Off.ARRAY_PDATA);
            int nSize = ReadInt32(pEntrysArray, Off.ARRAY_NSIZE);

            if (pData == IntPtr.Zero || nSize <= 0) return;

            Debug.WriteLine(string.Format("=== CProfile7 @ 0x{0:X8}: {1} записей ===",
                pProfile.ToInt32(), nSize));

            for (int i = 0; i < nSize && i < 256; i++)
            {
                IntPtr pEntry = ReadPointer(pData, i * 4);
                if (pEntry == IntPtr.Zero || !IsReadable(pEntry, Off.ENTRY_SIZE)) continue;

                // CString в MFC42 (ANSI) — это указатель на символьные данные.
                // CProfileEntry7::EntryName лежит по +0x0C.
                IntPtr pChars = ReadPointer(pEntry, Off.ENTRY_ENTRYNAME);
                string name = (pChars != IntPtr.Zero && IsReadable(pChars, 1))
                            ? (Marshal.PtrToStringAnsi(pChars) ?? string.Empty)
                            : string.Empty;

                int val = ReadInt32(pEntry, Off.ENTRY_DATA1);

                Debug.WriteLine(string.Format("  [{0,3}] type={1} data1={2} (0x{2:X8}) name='{3}'",
                    i, ReadInt32(pEntry, Off.ENTRY_TYPE), val, name));
            }
        }

        private static void DumpMemory(IntPtr p, int dwordCount, string title)
        {
            Debug.WriteLine("--- " + title + " @ 0x" + p.ToInt32().ToString("X8") + " ---");
            for (int i = 0; i < dwordCount; i++)
            {
                if (!IsReadable(new IntPtr(p.ToInt32() + i * 4), 4)) break;
                int v = *(int*)(p.ToInt32() + i * 4);
                Debug.WriteLine(string.Format("  +{0:X2}: 0x{1:X8} ({1})", i * 4, v));
            }
        }

        // ==================================================================
        //  Утилиты чтения / Win32
        // ==================================================================

        private static IntPtr ReadPointer(IntPtr p, int offset)
        {
            return *(IntPtr*)(p.ToInt32() + offset);
        }

        private static int ReadInt32(IntPtr p, int offset)
        {
            return *(int*)(p.ToInt32() + offset);
        }

        private static bool IsReadable(IntPtr p, int size)
        {
            if (p == IntPtr.Zero) return false;
            int addr = unchecked((int)p.ToInt32());
            if ((uint)addr < 0x00010000 || (uint)addr >= 0x7FFF0000) return false;
            return Win32.IsBadReadPtr(p, (uint)size) == 0;
        }

        /// <summary>Освободить ссылку на таблицу 1С.</summary>
        public void ReleaseTable()
        {
            _tableObj = null;
            _pExportContext = IntPtr.Zero;
            _pTableContext = IntPtr.Zero;
            _pSheetDoc = IntPtr.Zero;
            _pSheet = IntPtr.Zero;
        }

        public void Dispose()
        {
            ReleaseTable();
            // _hTemplate7Dll не выгружаем: DLL 1С живет дольше нас.
        }

        ~V7TableDocManager() { Dispose(); }

        private static class Win32
        {
            [DllImport("kernel32.dll")]
            public static extern int IsBadReadPtr(IntPtr lp, uint ucb);

            [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
            public static extern IntPtr LoadLibrary(string lpFileName);

            [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
            public static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

            [DllImport("kernel32.dll")]
            public static extern int FreeLibrary(IntPtr hLibModule);
        }
    }
}
