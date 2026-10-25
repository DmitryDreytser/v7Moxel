// PrintWrapper.cs
// Перехват РОДНОЙ печати 1С 7.7 и запись страниц в расширенные метафайлы (EMF) с последующей сборкой PDF.
//
// ИДЕЯ. Печатную форму рисует сама 1С (moxel.dll, CSheetView : CView) через plain GDI (CDC).
// Вместо перехвата каждого GDI-вызова подменяем КОНТЕКСТ: в момент StartDoc создаём контекст
// расширенного метафайла CreateEnhMetaFile(реальный_принтер_DC, frame = весь лист) и записываем
// его в CDC::m_hDC. Windows САМА протоколирует все GDI-вызовы в EMF; метрики устройства остаются
// принтерскими (EMF DC наследует capabilities reference DC) — пагинация 1С не меняется.
//
// РЕАЛЬНОЙ ПЕЧАТИ НЕТ: StartDoc/StartPage/EndPage/EndDoc перехватываются на уровне IAT
// (mfc42.dll + moxel.dll) и в спулер не уходят — бумага не расходуется.
//
// СХЕМА ПЕРЕХВАТА (x86, тот же стиль патчей, что и SaveWrapper):
//   1. CSheetView::OnBeginPrinting(pDC, pInfo) — prologue-патч 6 байт (PUSH addr; RET).
//      Хук запоминает CDC* и m_hDC (реальный принтер); оригинал вызывается через временное
//      снятие патча (restore-call-repatch; печать 1С однопоточная — безопасно).
//   2. StartDoc(hdc==hPrinter)  → создаём EMF DC страницы №1, пишем его в CDC::m_hDC, возврат «успех».
//   3. StartPage(hEmf)          → «успех» (запись продолжается).
//   4. EndPage(hEmf)            → страница готова: CloseEnhMetaFile → байты → создаём СВЕЖИЙ EMF DC
//                                  (чистые атрибуты = имитация сброса DC при реальном StartPage).
//   5. EndDoc(hEmf)             → закрываем последнюю страницу, возвращаем CDC::m_hDC = принтер.
//   6. CSheetView::OnEndPrinting(pDC, pInfo) → страницы → EmfPdfWriter.Save → снять все патчи.
//
// ЗАПУСК ПЕЧАТИ (два способа):
//   A) BeginPdfCapture(file); затем родная печать: Напечатать(1) через RCW (V7Table.InvokeTableMethod).
//      PDF готовится в OnEndPrinting, т.е. к возврату из Напечатать файл уже записан.
//   B) PrintWrapper.CaptureToPdf(pSheetDoc, file) — сам находит CView документа
//      (CDocument::GetFirstViewPosition/GetNextView) и вызывает CView::OnFilePrint (экспорт mfc42).
//
// Адреса НЕ зашиты: экспорты ищутся по подстроке имени в таблице экспорта — декорации не важны.
//
// Ограничения v1 (осознанно):
//   - PDF растровый (EmfPdfWriter) — текст не выделяется. Векторный EMF→PDF — v2.
//   - Печать должна идти через CView-цикл MFC (у открытой таблицы — так и есть). У никогда не
//     показанной таблицы view нет — вызывающий код должен дать fallback (PDFWriter/Pechkin).
//   - m_hAttribDC остаётся принтерским: MFC CDC при m_hDC != m_hAttribDC выбирает атрибуты в ОБА
//     контекста, поэтому GetTextMetrics/GetTextExtent дают ПЕЧАТНЫЕ метрики — это и нужно.
//   - Требуется принтер по умолчанию (метрики берутся с реального DC; самой печати не будет).

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media.Media3D;
using static Moxel.PrintWrapper;

namespace Moxel
{
    public static class PrintWrapper
    {

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT 
        {
            public int left, top, right, bottom;
            public override string ToString() => $"{left}, {top}, {right}, {bottom}";
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SIZE 
        {
            public int width, height;
            public override string ToString() => $"{width}, {height}";
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x, y;
            public override string ToString() => $"{x}, {y}";
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct ENHMETAHEADER
        {
            public uint iType, nSize;
            public RECT rclBounds, rclFrame;      // по 4 int
            public uint dSignature, nVersion, nBytes, nRecords;
            public ushort nHandles, sReserved;
            public uint nDescription, offDescription, nPalEntries;
            public SIZE szlDevice, szlMillimeters; // по 2 int
        }

        /// <summary>CPrintInfo из MFC 4.2 (afxpriv.h), 32-битная раскладка.</summary>
        [StructLayout(LayoutKind.Sequential, Size = 0x3C)]
        public struct CPrintInfo
        {
            public IntPtr m_pPD;                 // CPrintDialog* — pointer to native CPrintDialog

            [MarshalAs(UnmanagedType.Bool)]
            public bool m_bDocObject;           // BOOL

            [MarshalAs(UnmanagedType.Bool)]
            public bool m_bPreview;             // BOOL

            [MarshalAs(UnmanagedType.Bool)]
            public bool m_bDirect;              // BOOL

            [MarshalAs(UnmanagedType.Bool)]
            public bool m_bContinuePrinting;    // BOOL

            public uint m_nCurPage;             // UINT
            public uint m_nNumPreviewPages;     // UINT

            public IntPtr m_strPageDesc;        // CString by value — represent as pointer to characters (LPCTSTR)

            public IntPtr m_lpUserData;         // LPVOID

            public RECT m_rectDraw;             // CRect (= RECT)

            public uint m_nOffsetPage;          // UINT
            public uint m_dwFlags;              // DWORD
        }

        #region Константы раскладки

        // CDC (MFC42, x86): vft@0x00, HDC m_hDC@0x04, HDC m_hAttribDC@0x08, BOOL m_bPrinting@0x0C
        private const int CDC_OFF_HDC = 0x04;
        private const int CDC_SCAN_LIMIT = 0x20; // страховка раскладки: поиск m_hDC в первых 0x20 байтах

        // Индексы GetDeviceCaps для wingdi.h
        private const int CAP_TECHNOLOGY = 2;
        private const int CAP_HORZRES = 8;
        private const int CAP_VERTRES = 10;
        private const int CAP_LOGPIXELSX = 88;
        private const int CAP_LOGPIXELSY = 90;
        private const int CAP_PHYSICALWIDTH = 110;
        private const int CAP_PHYSICALHEIGHT = 111;
        private const int CAP_PHYSICALOFFSETX = 112;
        private const int CAP_PHYSICALOFFSETY = 113;
        private const int DT_RASPRINTER = 2;

        // Настройки виртуального принтера (A4, 600 DPI)
        private const int FB_DPI = 600;
        private const double FB_W_MM = 190.0; // Печатаемая область (с полями 10мм)
        private const double FB_H_MM = 277.0;
        private const double FB_PH_W_MM = 210.0; // Физический размер листа
        private const double FB_PH_H_MM = 297.0;

        private const string MoxModule = "Moxel.dll";
        private const string MfcModule = "mfc42.dll";

        private static readonly string[] HookedGdi = { "StartDocA", "StartDocW", "StartPage", "EndPage", "EndDoc", "GetDeviceCaps" };

        #endregion

        #region Делегаты нативных функций

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        private delegate void dPrintHook(IntPtr pView, IntPtr pDC, ref CPrintInfo pInfo);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int dStartDoc(IntPtr hdc, IntPtr pDocInfo);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int dOneArg(IntPtr hdc);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        private delegate IntPtr dGetFirstViewPosition(IntPtr pDoc);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        private delegate IntPtr dGetNextView(IntPtr pDoc, ref IntPtr pos);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        private delegate void dOnFilePrint(IntPtr pView);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int dGetDeviceCaps(IntPtr hdc, int index);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        private delegate int dDoPrintDialog(IntPtr pThis, IntPtr pPD);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        private delegate IntPtr dCreatePrinterDC(IntPtr pThis);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        private delegate IntPtr dCreateCPrintInfo(IntPtr pThis);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        delegate int dGetPrinterDeviceDefaults(IntPtr pApp, IntPtr pPD);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
        delegate int dGetStringProp(IntPtr thisCProfile, int nProp); // фактически возвращает CString* в eax

        #endregion

        #region Состояние

        private sealed class Capture
        {
            public string PdfFile;
            public EmfPdfSettings Settings;
            /// <summary>true — страницы отдавать EmfPdfWriter (растр), иначе EmfVectorPdfWriter (вектор).</summary>
            public bool Raster;
            public IntPtr pCDC;        // CDC* из OnBeginPrinting
            public IntPtr hPrinterDC;  // реальный DC принтера
            public IntPtr hEmfDC;      // текущий DC метафайла (страница в записи)
            public IntPtr FallbackDC; //  флаг, что мы используем фейковый DC
            public readonly List<byte[]> Pages = new List<byte[]>();
            public volatile bool Done;
            internal string PreferPrinter;
            internal string PrinterName;
            internal bool HasValidPD;
            internal IntPtr RefPrinterDC;
            internal IntPtr View;
        }

        private static Capture S; // печать 1С однопоточная — static достаточно

        private static bool _resolved;
        private static string _lastError = string.Empty;
        private static dCreateCPrintInfo CPrintInfo__ctor;
        private static dCreateCPrintInfo DestructCPrintInfo;
        static dGetPrinterDeviceDefaults _origGetPrinterDeviceDefaults;
        static dGetPrinterDeviceDefaults _hookGetPrinterDeviceDefaults;
        static IntPtr _hHoldMode, _hHoldNames; // копии на процесс, не освобождаем в цикле

        /// <summary>Число страниц последнего захвата.</summary>
        public static int LastPageCount { get; private set; }

        /// <summary>Текст последней ошибки (для диагностики на живой 1С).</summary>
        public static string LastError { get { return _lastError; } }

        /// <summary>Идёт ли захват (между BeginPdfCapture и финализацией).</summary>
        public static bool IsCapturing { get { return S != null; } }

        /// <summary>Вызывается после закрытия каждой страницы (аргумент — её номер, с 1).</summary>
        public static event Action<int> onPageCompleted;

        #endregion

        #region Адреса (поиск по таблице экспорта, без зашитых RVAs)

        private static IntPtr pOnBeginPrinting, pOnEndPrinting, pOnFilePrint, pGetFirstViewPosition, pGetNextView, pDoPrintDialog, pCreatePrinterDC, pOnPrint, pOnPrepareDC, pOnPreparePrinting;
        private static IntPtr pCPrintInfo;

        private static byte[] _savedBegin = new byte[6];
        private static byte[] _savedEnd = new byte[6];
        private static byte[] _savedDoPrintDialog = new byte[6];
        private static byte[] _savedCreatePrinterDC = new byte[6];

        private static byte[] _savedOnPrint1 = new byte[2];
        private static byte[] _savedOnPrint2 = new byte[2];

        private static dPrintHook _cbBegin, _cbEnd;
        private static dPrintHook _origOnBegin, _origOnEnd;
        private static dStartDoc _cbStartDocA, _cbStartDocW, _origStartDocA, _origStartDocW;
        private static dOneArg _cbStartPage, _cbEndPage, _cbEndDoc, _origStartPage, _origEndPage, _origEndDoc;
        private static dGetFirstViewPosition _origGetFirstViewPosition;
        private static dGetNextView _origGetNextView;
        private static dOnFilePrint _origOnFilePrint;
        private static dGetDeviceCaps _cbGetDeviceCaps, _origGetDeviceCaps;
        private static dPrintHook _origOnPrint;
        private static dDoPrintDialog _cbDoPrintDialog;
        private static dCreatePrinterDC _cbCreatePrinterDC, _origCreatePrinterDC;


        private sealed class IatSlot { public IntPtr SlotAddr; public IntPtr Orig; }
        private static readonly List<IatSlot> _iatPatched = new List<IatSlot>();

        /// <summary>Доступен ли перехват (экспорты найдены). Безопасно вызывать всегда.</summary>
        public static bool IsAvailable
        {
            get { Resolve(); return pOnBeginPrinting != IntPtr.Zero && pOnEndPrinting != IntPtr.Zero; }
        }

        private static void Resolve()
        {
            if (_resolved) return;
            _resolved = true;

            try
            {
                IntPtr hMoxel = GetModuleHandle(MoxModule);
                IntPtr hMfc = GetModuleHandle(MfcModule);

                if (hMoxel != IntPtr.Zero)
                {
                    pOnBeginPrinting = WinApi.GetProcAddress(hMoxel, "?OnBeginPrinting@CSheetView@@MAEXPAVCDC@@PAUCPrintInfo@@@Z");// FindExport(hMoxel, "OnBeginPrinting@CSheetView");
                    pOnEndPrinting = WinApi.GetProcAddress(hMoxel, "?OnEndPrinting@CSheetView@@MAEXPAVCDC@@PAUCPrintInfo@@@Z");// FindExport(hMoxel, "OnEndPrinting@CSheetView");
                    pOnPrint = WinApi.GetProcAddress(hMoxel, "?OnPrint@CSheetView@@MAEXPAVCDC@@PAUCPrintInfo@@@Z");
                    pOnPrepareDC = WinApi.GetProcAddress(hMoxel, "?OnPrepareDC@CSheetView@@UAEXPAVCDC@@PAUCPrintInfo@@@Z");
                    pOnPreparePrinting = WinApi.GetProcAddress(hMoxel, "?OnPreparePrinting@CSheetView@@MAEHPAUCPrintInfo@@@Z");
                }

                if (hMfc != IntPtr.Zero)
                {
                    pOnFilePrint = WinApi.GetProcAddress(hMfc, 4613);
                    pGetFirstViewPosition = WinApi.GetProcAddress(hMfc, 3198);// "GetFirstViewPosition@CDocument");
                    pGetNextView = WinApi.GetProcAddress(hMfc, 3454);// "GetNextView@CDocument");
                    pDoPrintDialog = WinApi.GetProcAddress(hMfc, 2537); // "?DoPrintDialog@CWinApp@@QAEHPAVCPrintDialog@@@Z");
                    pCreatePrinterDC = WinApi.GetProcAddress(hMfc, 2246); // "?CreatePrinterDC@CPrintDialog@@QAEPAUHDC__@@XZ");
                }

                Debug.WriteLine(
                    $"PrintWrapper: OnBeginPrinting={pOnBeginPrinting.ToInt32():X8} OnEndPrinting={pOnEndPrinting.ToInt32():X8} " +
                    $"OnFilePrint={pOnFilePrint.ToInt32():X8} GetFirstViewPosition={pGetFirstViewPosition.ToInt32():X8} " +
                    $"GetNextView={pGetNextView.ToInt32():X8}");
            }
            catch (Exception ex)
            {
                _lastError = "Resolve: " + ex.Message;
                Debug.WriteLine("PrintWrapper.Resolve: " + ex);
            }
        }

        #endregion

        #region Публичный API

        /// <summary>
        /// Вооружить перехват. Родная печать (Напечатать(1)/кнопка печати) запишет страницы в PDF
        /// и финализируется в OnEndPrinting. Реальной печати не будет.
        /// </summary>
        public static bool BeginPdfCapture(string pdfFile, EmfPdfSettings settings = null, bool raster = false)
        {
            Resolve();
            _lastError = string.Empty;

            if (pOnBeginPrinting == IntPtr.Zero || pOnEndPrinting == IntPtr.Zero)
            {
                _lastError = "Экспорты CSheetView::OnBeginPrinting/OnEndPrinting не найдены";
                return false;
            }
            if (string.IsNullOrEmpty(pdfFile))
            {
                _lastError = "Не задано имя PDF-файла";
                return false;
            }
            if (S != null)
                EndPdfCapture(); // снять незавершённый захват

            try
            {
                // Оригиналы GDI-функций (для пропуска не-принтерных вызовов)
                IntPtr hGdi = LoadLibrary("gdi32.dll");
                if (hGdi == IntPtr.Zero)
                {
                    _lastError = "gdi32.dll недоступен";
                    return false;
                }

                _origStartDocA = (dStartDoc)Marshal.GetDelegateForFunctionPointer(GetProcAddress(hGdi, "StartDocA"), typeof(dStartDoc));
                _origStartDocW = (dStartDoc)Marshal.GetDelegateForFunctionPointer(GetProcAddress(hGdi, "StartDocW"), typeof(dStartDoc));
                _origStartPage = (dOneArg)Marshal.GetDelegateForFunctionPointer(GetProcAddress(hGdi, "StartPage"), typeof(dOneArg));
                _origEndPage = (dOneArg)Marshal.GetDelegateForFunctionPointer(GetProcAddress(hGdi, "EndPage"), typeof(dOneArg));
                _origEndDoc = (dOneArg)Marshal.GetDelegateForFunctionPointer(GetProcAddress(hGdi, "EndDoc"), typeof(dOneArg));
                _origOnBegin = (dPrintHook)Marshal.GetDelegateForFunctionPointer(pOnBeginPrinting, typeof(dPrintHook));
                _origOnEnd = (dPrintHook)Marshal.GetDelegateForFunctionPointer(pOnEndPrinting, typeof(dPrintHook));
                _origGetDeviceCaps = (dGetDeviceCaps)Marshal.GetDelegateForFunctionPointer(GetProcAddress(hGdi, "GetDeviceCaps"), typeof(dGetDeviceCaps));

                _origOnPrint = (dPrintHook)Marshal.GetDelegateForFunctionPointer(pOnPrint, typeof(dPrintHook));


                _cbStartDocA = StartDocAHook;
                _cbStartDocW = StartDocWHook;
                _cbStartPage = StartPageHook;
                _cbEndPage = EndPageHook;
                _cbEndDoc = EndDocHook;
                _cbBegin = OnBeginPrintingHook;
                _cbEnd = OnEndPrintingHook;
                _cbGetDeviceCaps = GetDeviceCapsHook;
                _cbDoPrintDialog = DoPrintDialogHook;

                S = new Capture { PdfFile = pdfFile, Settings = settings ?? new EmfPdfSettings(), Raster = raster };
                LastPageCount = 0;

                // 1) Процедурные патчи CSheetView (6 байт: PUSH stub; RET)
                if (!SaveAndPatch6(pOnBeginPrinting, _cbBegin, _savedBegin) ||
                    !SaveAndPatch6(pOnEndPrinting, _cbEnd, _savedEnd))
                    throw new InvalidOperationException("Не удалось пропатчить OnBeginPrinting/OnEndPrinting");

                if (pCreatePrinterDC != IntPtr.Zero)
                {
                    _origCreatePrinterDC = (dCreatePrinterDC)Marshal.GetDelegateForFunctionPointer(pCreatePrinterDC, typeof(dCreatePrinterDC));
                    _cbCreatePrinterDC = CreatePrinterDCHook;
                    if (!SaveAndPatch6(pCreatePrinterDC, _cbCreatePrinterDC, _savedCreatePrinterDC))
                        throw new InvalidOperationException("Не удалось пропатчить CreatePrinterDC");
                }

                if (pDoPrintDialog != IntPtr.Zero)
                {
                    if (!SaveAndPatch6(pDoPrintDialog, _cbDoPrintDialog, _savedDoPrintDialog))
                        throw new InvalidOperationException("Не удалось пропатчить DoPrintDialog");
                }


                // 2) IAT moxel.dll + mfc42.dll: StartDocA/W, StartPage, EndPage, EndDoc → заглушки
                PatchIat(GetModuleHandle(MoxModule));
                PatchIat(GetModuleHandle(MfcModule));

                PatchOnPrint();

                Debug.WriteLine($"PrintWrapper: захват вооружён → {pdfFile}");
                return true;
            }
            catch (Exception ex)
            {
                _lastError = "BeginPdfCapture: " + ex.Message;
                Debug.WriteLine(_lastError);
                UnpatchAll();
                S = null;
                return false;
            }
        }

        private static int DoPrintDialogHook(IntPtr pThis, IntPtr pPD)
        {
            // Глушим диалог, имитируя успешное нажатие "ОК"
            Debug.WriteLine("PrintWrapper: DoPrintDialog bypassed, returning 1 (IDOK)");
            return 1;
        }

        private static int GetDeviceCapsHook(IntPtr hdc, int index)
        {
            // Если это наш фейковый DC (нет системного принтера), отдаем метрики A4 600 DPI
            if (S != null && !S.Done && S.FallbackDC != IntPtr.Zero && hdc == S.FallbackDC)
            {
                switch (index)
                {
                    case CAP_TECHNOLOGY: return DT_RASPRINTER; // Говорим, что мы принтер, а не экран
                    case CAP_LOGPIXELSX: return FB_DPI;
                    case CAP_LOGPIXELSY: return FB_DPI;
                    case CAP_HORZRES: return (int)(FB_W_MM / 25.4 * FB_DPI);
                    case CAP_VERTRES: return (int)(FB_H_MM / 25.4 * FB_DPI);
                    case CAP_PHYSICALWIDTH: return (int)(FB_PH_W_MM / 25.4 * FB_DPI);
                    case CAP_PHYSICALHEIGHT: return (int)(FB_PH_H_MM / 25.4 * FB_DPI);
                    case CAP_PHYSICALOFFSETX: return (int)(10.0 / 25.4 * FB_DPI); // Поля 10мм
                    case CAP_PHYSICALOFFSETY: return (int)(10.0 / 25.4 * FB_DPI);
                }
            }
            return _origGetDeviceCaps(hdc, index);
        }

        /// <summary>
        /// Принудительная финализация (если OnEndPrinting не наступил).
        /// Возвращает число записанных страниц (PDF готов) или -1.
        /// </summary>
        public static int EndPdfCapture()
        {
            if (S == null)
                return LastPageCount > 0 ? LastPageCount : -1;

            try { return FinalizeCapture(); }
            finally { S = null; }
        }

        /// <summary>
        /// Высокий уровень: PDF одной таблицы. Сам находит CView документа и запускает родной
        /// цикл печати (CView::OnFilePrint). true — PDF записан; иначе caller должен дать fallback.
        /// </summary>
        public static bool CaptureToPdf(IntPtr pSheetDoc, string pdfFile, EmfPdfSettings settings = null, V7Table table = null, bool raster = false)
        {
            Resolve();
            if (pSheetDoc == IntPtr.Zero || string.IsNullOrEmpty(pdfFile))
                return false;

            if (pOnFilePrint == IntPtr.Zero || pGetFirstViewPosition == IntPtr.Zero || pGetNextView == IntPtr.Zero)
            {
                _lastError = "Экспорты mfc42 (OnFilePrint/GetFirstViewPosition/GetNextView) не найдены";
                return false;
            }

            if (!BeginPdfCapture(pdfFile, settings, raster))
                return false;

            try
            {
                _origGetFirstViewPosition = (dGetFirstViewPosition)Marshal.GetDelegateForFunctionPointer(
                    pGetFirstViewPosition, typeof(dGetFirstViewPosition));
                _origGetNextView = (dGetNextView)Marshal.GetDelegateForFunctionPointer(
                    pGetNextView, typeof(dGetNextView));

                IntPtr pos = _origGetFirstViewPosition(pSheetDoc);
                IntPtr view = _origGetNextView(pSheetDoc, ref pos);
                if (view == IntPtr.Zero && table == null)
                {
                    _lastError = "У документа нет CView (таблица не показывалась) — перехват печати невозможен";
                    Debug.WriteLine("PrintWrapper: " + _lastError);
                    EndPdfCapture();
                    return false;
                }

                _origOnFilePrint = (dOnFilePrint)Marshal.GetDelegateForFunctionPointer(pOnFilePrint, typeof(dOnFilePrint));
                if (view != IntPtr.Zero)
                    _origOnFilePrint(view); // синхронный цикл MFC: StartDoc → {StartPage, OnPrepareDC, OnPrint, EndPage}* → EndDoc
                else if (table != null)
                    table.Print(0);

                if (S != null)
                    EndPdfCapture(); // OnEndPrinting уже отработал бы; страховка

                return LastPageCount > 0 && System.IO.File.Exists(pdfFile);
            }
            catch (Exception ex)
            {
                _lastError = "CaptureToPdf: " + ex.Message;
                Debug.WriteLine(_lastError);
                try { EndPdfCapture(); } catch { }
                return false;
            }
        }

        #endregion

        #region Патч порядка вывода картинок и текста
        private static void PatchOnPrint()
        {
            if (pOnPrint == IntPtr.Zero) return;

            IntPtr ptr1 = new IntPtr(pOnPrint.ToInt32() + 0x557);
            IntPtr ptr2 = new IntPtr(pOnPrint.ToInt32() + 0x565);

            // Сохраняем оригинальные байты (2 байта на каждый адрес)
            Marshal.Copy(ptr1, _savedOnPrint1, 0, 2);
            Marshal.Copy(ptr2, _savedOnPrint2, 0, 2);

            // Проверяем, что оригинальные байты совпадают с ожидаемыми (1AC5 и 4127).
            // Если версия moxel.dll другая и байты не совпадают — патч не ставим, чтобы не уронить 1С!
            short orig1 = Marshal.ReadInt16(ptr1);
            short orig2 = Marshal.ReadInt16(ptr2);

            if (orig1 != 0x1AC5 || orig2 != 0x4127)
            {
                Debug.WriteLine($"PrintWrapper: OnPrint patch skipped (version mismatch). Expected 1AC5/4127, got {orig1:X}/{orig2:X}");
                return;
            }

            uint oldProtect;
            // Патчим первый вызов: call CSheetView::DrawObjects
            if (VirtualProtect(ptr1, 2, 0x40 /*PAGE_EXECUTE_READWRITE*/, out oldProtect))
            {
                Marshal.WriteInt16(ptr1, 0x4135); // Запись 2 байт (35 41)
                uint tmp; VirtualProtect(ptr1, 2, oldProtect, out tmp);
            }

            // Патчим второй вызов: call CSheetView::DrawSheet
            if (VirtualProtect(ptr2, 2, 0x40, out oldProtect))
            {
                Marshal.WriteInt16(ptr2, 0x1AB7); // Запись 2 байт (B7 1A)
                uint tmp; VirtualProtect(ptr2, 2, oldProtect, out tmp);
            }

            FlushInstructionCache(GetCurrentProcess(), pOnPrint, (UIntPtr)0x600);
            Debug.WriteLine("PrintWrapper: OnPrint patched (DrawObjects/DrawSheet swapped)");
        }

        private static void UnpatchOnPrint()
        {
            if (pOnPrint == IntPtr.Zero) return;
            // Если _savedOnPrint1[0] == 0, значит патч не был применен (сработала защита версии)
            if (_savedOnPrint1[0] == 0 && _savedOnPrint1[1] == 0) return;

            IntPtr ptr1 = new IntPtr(pOnPrint.ToInt32() + 0x557);
            IntPtr ptr2 = new IntPtr(pOnPrint.ToInt32() + 0x565);

            uint oldProtect;
            if (VirtualProtect(ptr1, 2, 0x40, out oldProtect))
            {
                Marshal.Copy(_savedOnPrint1, 0, ptr1, 2);
                uint tmp; VirtualProtect(ptr1, 2, oldProtect, out tmp);
            }

            if (VirtualProtect(ptr2, 2, 0x40, out oldProtect))
            {
                Marshal.Copy(_savedOnPrint2, 0, ptr2, 2);
                uint tmp; VirtualProtect(ptr2, 2, oldProtect, out tmp);
            }

            FlushInstructionCache(GetCurrentProcess(), pOnPrint, (UIntPtr)0x600);
        }

        #endregion

        #region Хуки печати
     
        private static void OnBeginPrintingHook(IntPtr pView, IntPtr pDC, ref CPrintInfo pInfo)
        {
            bool active = S != null;
            if (active)
            {
                S.pCDC = pDC;
                S.hPrinterDC = ReadHdc(pDC);
                Debug.WriteLine($"PrintWrapper: OnBeginPrinting CDC={pDC.ToInt32():X8} hDC={S.hPrinterDC.ToInt32():X8}");
            }

            // Оригинал: снять патч → вызвать → вернуть патч (печать однопоточная)
            Unpatch6(pOnBeginPrinting, _savedBegin);
            try
            {
                _origOnBegin(pView, pDC, ref pInfo);
            }
            finally
            {
                if (S != null)
                    Repatch6(pOnBeginPrinting, _cbBegin);
            }
        }

        private static void OnEndPrintingHook(IntPtr pView, IntPtr pDC, ref CPrintInfo pInfo)
        {

            Unpatch6(pOnEndPrinting, _savedEnd);
            try
            {
                _origOnEnd(pView, pDC, ref pInfo);
            }
            finally
            {
                if (S != null && !S.Done)
                    Repatch6(pOnEndPrinting, _cbEnd);
            }

            if (S != null)
                FinalizeCapture(); // сбор страниц → PDF, снятие всех патчей
        }

        private static int StartDocAHook(IntPtr hdc, IntPtr pDocInfo) { return StartDocCore(hdc, _origStartDocA, pDocInfo); }
        private static int StartDocWHook(IntPtr hdc, IntPtr pDocInfo) { return StartDocCore(hdc, _origStartDocW, pDocInfo); }

        private static int StartDocCore(IntPtr hdc, dStartDoc orig, IntPtr pDocInfo)
        {
            if (S != null && !S.Done && S.hEmfDC == IntPtr.Zero &&
                S.pCDC != IntPtr.Zero && hdc == S.hPrinterDC)
            {
                S.hEmfDC = CreatePageEmf();
                if (S.hEmfDC != IntPtr.Zero)
                {
                    Marshal.WriteIntPtr(S.pCDC, CDC_OFF_HDC, S.hEmfDC);
                    Debug.WriteLine($"PrintWrapper: StartDoc → EMF DC {S.hEmfDC.ToInt32():X8} (страница 1)");
                    return 1; // успех, в спулер не уходим
                }
            }
            return orig(hdc, pDocInfo);
        }

        private static int StartPageHook(IntPtr hdc)
        {
            if (S != null && !S.Done && hdc == S.hEmfDC)
                return 1; // запись уже идёт
            return _origStartPage(hdc);
        }

        private static int EndPageHook(IntPtr hdc)
        {
            if (S != null && !S.Done && hdc == S.hEmfDC)
            {
                byte[] page = TakeEmfBytes();
                if (page != null)
                {
                    S.Pages.Add(page);
                    var h = onPageCompleted;
                    if (h != null) h(S.Pages.Count);
                }

                // Свежий EMF DC = чистые атрибуты (имитация сброса DC при реальном StartPage)
                S.hEmfDC = CreatePageEmf();
                if (S.hEmfDC != IntPtr.Zero && S.pCDC != IntPtr.Zero)
                    Marshal.WriteIntPtr(S.pCDC, CDC_OFF_HDC, S.hEmfDC);

                Debug.WriteLine($"PrintWrapper: EndPage → страница {S.Pages.Count} готова ({(page != null ? page.Length : 0)} байт)");
                return 1;
            }
            return _origEndPage(hdc);
        }

        private static int EndDocHook(IntPtr hdc)
        {
            if (S != null && !S.Done && hdc == S.hEmfDC)
            {
                byte[] page = TakeEmfBytes();
                if (page != null)
                {
                    S.Pages.Add(page);
                    var h = onPageCompleted;
                    if (h != null) h(S.Pages.Count);
                }

                if (S.pCDC != IntPtr.Zero)
                    Marshal.WriteIntPtr(S.pCDC, CDC_OFF_HDC, S.hPrinterDC); // вернуть как было

                S.hEmfDC = IntPtr.Zero;
                S.Done = true;
                return 1;
            }
            return _origEndDoc(hdc);
        }

        private static IntPtr CreatePrinterDCHook(IntPtr pThis)
        {
            // Сначала вызываем оригинал. Если есть системный принтер — он вернет HDC.
            Unpatch6(pCreatePrinterDC, _savedCreatePrinterDC);
            IntPtr hdc;
            try 
            { 
                hdc = _origCreatePrinterDC(pThis); 
            }
            finally 
            { 
                if (S != null) 
                    Repatch6(pCreatePrinterDC, _cbCreatePrinterDC); 
            }

            // Если системного принтера нет (hdc == 0) — создаем фейковый A4 600 DPI
            if (hdc == IntPtr.Zero && S != null)
            {
                hdc = CreateICA("DISPLAY", null, null, IntPtr.Zero);
                S.FallbackDC = hdc; // Помечаем для хука GetDeviceCaps

                // КРИТИЧЕСКИ ВАЖНО: Записываем фейковый DC в m_pd.hdc (смещение 0x70)
                // В MFC 4.2 PRINTDLG начинается со смещения 0x60, поле hdc в ней — на +0x10.
                // Таким образом, 0x60 + 0x10 = 0x70.
                Marshal.WriteIntPtr(pThis, 0x70, hdc);

                Debug.WriteLine("PrintWrapper: CreatePrinterDC failed. Injected A4 600 DPI fallback DC.");
            }
            return hdc;
        }

        #endregion

        #region Служебные операции захвата

        private static IntPtr ReadHdc(IntPtr pCDC)
        {
            IntPtr h = Marshal.ReadIntPtr(pCDC, CDC_OFF_HDC);
            if (h == IntPtr.Zero)
            {
                for (int off = 4; off <= CDC_SCAN_LIMIT; off += 4)
                {
                    h = Marshal.ReadIntPtr(pCDC, off);
                    if (h != IntPtr.Zero) break;
                }
            }
            return h;
        }

        /// <summary>EMF DC с reference DC принтера и frame'ом на весь лист (0.01 мм).</summary>
        private static IntPtr CreatePageEmf()
        {
            RECT frame = ComputeFrameRect(S.hPrinterDC);
            return CreateEnhMetaFileA(S.hPrinterDC, IntPtr.Zero, ref frame, IntPtr.Zero);
        }

        private static RECT ComputeFrameRect(IntPtr hPrinterDC)
        {
            int w, h, dpiX, dpiY;

            if (S.FallbackDC != IntPtr.Zero)
            {
                // Используем хардкод для A4 600 DPI
                w = (int)(FB_W_MM / 25.4 * FB_DPI);
                h = (int)(FB_H_MM / 25.4 * FB_DPI);
                dpiX = FB_DPI;
                dpiY = FB_DPI;
            }
            else
            {
                // Стандартный запрос к реальному принтеру
                w = GetDeviceCaps(hPrinterDC, CAP_HORZRES);
                h = GetDeviceCaps(hPrinterDC, CAP_VERTRES);
                dpiX = Math.Max(1, GetDeviceCaps(hPrinterDC, CAP_LOGPIXELSX));
                dpiY = Math.Max(1, GetDeviceCaps(hPrinterDC, CAP_LOGPIXELSY));
            }

            return new RECT
            {
                left = 0,
                top = 0,
                right = Math.Max(1, w * 2540 / dpiX - 1),  // 0.01 мм
                bottom = Math.Max(1, h * 2540 / dpiY - 1)
            };
        }

        /// <summary>Закрыть текущий EMF DC и вернуть его байты (null — страница пустая).</summary>
        private static byte[] TakeEmfBytes()
        {
            if (S.hEmfDC == IntPtr.Zero)
                return null;

            IntPtr hmf = CloseEnhMetaFile(S.hEmfDC);
            S.hEmfDC = IntPtr.Zero;
            if (hmf == IntPtr.Zero)
                return null;

            try
            {
                //byte[] hdr = new byte[88];
                if (GetEnhMetaFileHeader(hmf, 88, out var hdr) == 0)
                    return null;

                int bl = hdr.rclBounds.left, bt = hdr.rclBounds.top, br = hdr.rclBounds.right, bb = hdr.rclBounds.bottom;
                if (br <= bl || bb <= bt)
                    return null; // пустая страница (хвост ротации)

                uint size = GetEnhMetaFileBits(hmf, 0, null);
                if (size == 0)
                    return null;

                byte[] data = new byte[size];
                if (GetEnhMetaFileBits(hmf, size, data) != size)
                    return null;

                return data;
            }
            finally
            {
                DeleteEnhMetaFile(hmf);
            }
        }

        /// <summary>Собрать страницы → PDF, снять все патчи. Возвращает число страниц.</summary>
        private static int FinalizeCapture()
        {
            Capture cap = S;
            if (cap == null)
                return -1;

            // Незакрытая страница (печать оборвалась до EndDoc)
            if (cap.hEmfDC != IntPtr.Zero)
            {
                byte[] page = TakeEmfBytes();
                if (page != null)
                    cap.Pages.Add(page);
                if (cap.pCDC != IntPtr.Zero)
                    Marshal.WriteIntPtr(cap.pCDC, CDC_OFF_HDC, cap.hPrinterDC);
            }

            UnpatchAll();
            S = null;

            LastPageCount = cap.Pages.Count;
            if (cap.Pages.Count == 0)
            {
                _lastError = "Страниц не получено — печать не прошла через перехваченный контур " +
                             "(нет view/принтера либо иной путь печати)";
                Debug.WriteLine("PrintWrapper: " + _lastError);
                return -1;
            }

            // Выбор писателя EMF -> PDF: вектор (EmfVectorPdfWriter, текст выделяется)
            // или растр (EmfPdfWriter, растеризация с настройками Dpi/JPEG).
            Func<IList<byte[]>, string, EmfPdfSettings, bool> writer =
                cap.Raster ? (Func<IList<byte[]>, string, EmfPdfSettings, bool>)EmfPdfWriter.Save
                           : EmfVectorPdfWriter.Save;
            bool ok = writer(cap.Pages, cap.PdfFile, cap.Settings);
            Debug.WriteLine($"PrintWrapper: {(cap.Raster ? "растровый" : "векторный")} PDF {(ok ? "записан" : "НЕ записан")}, страниц {cap.Pages.Count} → {cap.PdfFile}");
            return ok ? cap.Pages.Count : -1;
        }

        private static void UnpatchAll()
        {
            try
            {
                UnpatchOnPrint();
                Unpatch6(pOnBeginPrinting, _savedBegin);
                Unpatch6(pOnEndPrinting, _savedEnd);
                if (pDoPrintDialog != IntPtr.Zero)
                    Unpatch6(pDoPrintDialog, _savedDoPrintDialog);
                if (pCreatePrinterDC != IntPtr.Zero)
                    Unpatch6(pCreatePrinterDC, _savedCreatePrinterDC);
            }
            catch (Exception ex) { Debug.WriteLine("PrintWrapper.UnpatchAll(procs): " + ex); }

            foreach (IatSlot slot in _iatPatched)
            {
                try { WriteIatSlot(slot.SlotAddr, slot.Orig); }
                catch (Exception ex) { Debug.WriteLine("PrintWrapper.UnpatchAll(iat): " + ex); }
            }
            _iatPatched.Clear();
        }

        #endregion

        #region Процедурные патчи 6 байт (PUSH imm32; RET)

        private static bool SaveAndPatch6(IntPtr addr, Delegate stub, byte[] saved)
        {
            IntPtr fn = Marshal.GetFunctionPointerForDelegate(stub);
            Marshal.Copy(addr, saved, 0, 6);

            uint oldProtect;
            if (!VirtualProtect(addr, 6, 0x40 /*PAGE_EXECUTE_READWRITE*/, out oldProtect))
                return false;
            try
            {
                Marshal.WriteByte(addr, 0x68); // PUSH
                Marshal.WriteIntPtr(new IntPtr(addr.ToInt32() + 1), fn);
                Marshal.WriteByte(new IntPtr(addr.ToInt32() + 5), 0xC3); // RET
                FlushInstructionCache(GetCurrentProcess(), addr, (UIntPtr)6);
                return true;
            }
            finally
            {
                uint tmp;
                VirtualProtect(addr, 6, oldProtect, out tmp);
            }
        }

        private static void Unpatch6(IntPtr addr, byte[] saved)
        {
            uint oldProtect;
            if (!VirtualProtect(addr, 6, 0x40, out oldProtect))
                return;
            try
            {
                Marshal.Copy(saved, 0, addr, 6);
                FlushInstructionCache(GetCurrentProcess(), addr, (UIntPtr)6);
            }
            finally
            {
                uint tmp2;
                VirtualProtect(addr, 6, oldProtect, out tmp2);
            }
        }

        private static void Repatch6(IntPtr addr, Delegate stub)
        {
            IntPtr fn = Marshal.GetFunctionPointerForDelegate(stub);
            uint oldProtect;
            if (!VirtualProtect(addr, 6, 0x40, out oldProtect))
                return;
            try
            {
                Marshal.WriteByte(addr, 0x68);
                Marshal.WriteIntPtr(new IntPtr(addr.ToInt32() + 1), fn);
                Marshal.WriteByte(new IntPtr(addr.ToInt32() + 5), 0xC3);
                FlushInstructionCache(GetCurrentProcess(), addr, (UIntPtr)6);
            }
            finally
            {
                uint tmp;
                VirtualProtect(addr, 6, oldProtect, out tmp);
            }
        }

        #endregion

        #region IAT-патчинг (mfc42.dll / moxel.dll → gdi32)

        private static void PatchIat(IntPtr hModule)
        {
            if (hModule == IntPtr.Zero)
                return;

            foreach (string func in HookedGdi)
            {
                IntPtr stub;
                switch (func)
                {
                    case "StartDocA": stub = Marshal.GetFunctionPointerForDelegate(_cbStartDocA); break;
                    case "StartDocW": stub = Marshal.GetFunctionPointerForDelegate(_cbStartDocW); break;
                    case "StartPage": stub = Marshal.GetFunctionPointerForDelegate(_cbStartPage); break;
                    case "EndPage": stub = Marshal.GetFunctionPointerForDelegate(_cbEndPage); break;
                    case "EndDoc": stub = Marshal.GetFunctionPointerForDelegate(_cbEndDoc); break;
                    case "GetDeviceCaps": stub = Marshal.GetFunctionPointerForDelegate(_cbGetDeviceCaps); break;
                    default: continue;
                }

                foreach (IntPtr slot in FindIatSlots(hModule, "gdi32.dll", func))
                {
                    IntPtr orig = Marshal.ReadIntPtr(slot);
                    if (orig == IntPtr.Zero || orig == stub)
                        continue;
                    if (WriteIatSlot(slot, stub))
                        _iatPatched.Add(new IatSlot { SlotAddr = slot, Orig = orig });
                }
            }
        }

        private static bool WriteIatSlot(IntPtr slot, IntPtr value)
        {
            uint oldProtect;
            if (!VirtualProtect(slot, 4, 0x40, out oldProtect))
                return false;
            try
            {
                Marshal.WriteIntPtr(slot, value);
                return true;
            }
            finally
            {
                uint tmp;
                VirtualProtect(slot, 4, oldProtect, out tmp);
            }
        }

        /// <summary>Все IAT-слоты модуля для указанной функции (обычно один).</summary>
        private static IEnumerable<IntPtr> FindIatSlots(IntPtr hModule, string dllName, string funcName)
        {
            int eLfanew = Marshal.ReadInt32(hModule, 0x3C);
            IntPtr opt = new IntPtr(hModule.ToInt32() + eLfanew + 0x18);      // PE32: OptionalHeader
            int importRva = Marshal.ReadInt32(new IntPtr(opt.ToInt32() + 0x68)); // DataDirectory[1] = Import
            if (importRva == 0)
                yield break;

            IntPtr desc = new IntPtr(hModule.ToInt32() + importRva);
            while (true)
            {
                int oft = Marshal.ReadInt32(desc, 0x00);        // OriginalFirstThunk
                int nameRva = Marshal.ReadInt32(desc, 0x0C);    // Name
                int ft = Marshal.ReadInt32(desc, 0x10);         // FirstThunk
                if (nameRva == 0 && ft == 0)
                    yield break;

                if (nameRva != 0 && oft != 0)
                {
                    string dll = Marshal.PtrToStringAnsi(new IntPtr(hModule.ToInt32() + nameRva));
                    if (string.Equals(dll, dllName, StringComparison.OrdinalIgnoreCase))
                    {
                        IntPtr thunk = new IntPtr(hModule.ToInt32() + oft);
                        int i = 0;
                        while (true)
                        {
                            int hintNameRva = Marshal.ReadInt32(thunk, i * 4);
                            if (hintNameRva == 0)
                                break;
                            if ((hintNameRva & unchecked((int)0x80000000)) == 0) // импорт по имени
                            {
                                IntPtr byName = new IntPtr(hModule.ToInt32() + hintNameRva);
                                string name = Marshal.PtrToStringAnsi(new IntPtr(byName.ToInt32() + 2)); // +2: пропустить WORD hint
                                if (string.Equals(name, funcName, StringComparison.Ordinal))
                                    yield return new IntPtr(hModule.ToInt32() + ft + i * 4);
                            }
                            i++;
                        }
                    }
                }
                desc = new IntPtr(desc.ToInt32() + 20); // IMAGE_IMPORT_DESCRIPTOR
            }
        }

        #endregion

        #region WinAPI

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool VirtualProtect(IntPtr lpAddress, int dwSize, uint flNewProtect, out uint lpflOldProtect);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FlushInstructionCache(IntPtr hProcess, IntPtr lpBaseAddress, UIntPtr dwSize);

        [DllImport("gdi32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr CreateEnhMetaFileA(IntPtr hdcRef, IntPtr lpDescription, ref RECT lprcFrame, IntPtr lpStrDescription);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CloseEnhMetaFile(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteEnhMetaFile(IntPtr hEmf);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetEnhMetaFileHeader(IntPtr hEmf, uint nSize, out ENHMETAHEADER hdr);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetEnhMetaFileBits(IntPtr hEmf, uint nSize, byte[] lpData);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern int GetDeviceCaps(IntPtr hdc, int index);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool GetPageMetricsA(IntPtr hdc, out POINT p);

        [DllImport("gdi32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr CreateICA(string lpszDriver, string lpszDevice, string lpszOutput, IntPtr lpInitData);

        #endregion
    }
}
