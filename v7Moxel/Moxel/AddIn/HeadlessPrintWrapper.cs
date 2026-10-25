using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OneC77.Headless
{
    #region Win32

    internal static class K
    {
        [DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern IntPtr GetModuleHandleA(string name);

        [DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern IntPtr GetProcAddress(IntPtr module, string name);

        [DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern IntPtr GetProcAddress(IntPtr module, int ordinal);

        [DllImport("kernel32")]
        public static extern IntPtr GlobalAlloc(uint flags, IntPtr bytes);

        [DllImport("kernel32")]
        public static extern IntPtr GlobalLock(IntPtr mem);

        [DllImport("kernel32")]
        public static extern int GlobalUnlock(IntPtr mem);

        [DllImport("kernel32")]
        public static extern IntPtr GlobalFree(IntPtr mem);

        [DllImport("kernel32")]
        public static extern int GlobalSize(IntPtr mem);

        [DllImport("kernel32")]
        public static extern int GlobalFlags(IntPtr mem);

        [DllImport("kernel32")]
        public static extern IntPtr VirtualAlloc(IntPtr addr, IntPtr size, uint type, uint protect);

        [DllImport("kernel32")]
        public static extern bool VirtualProtect(IntPtr addr, IntPtr size, uint prot, out uint old);
    }

    internal static class GDI
    {
        [DllImport("gdi32")]
        public static extern IntPtr CreateDC(string drv, string dev, string port, IntPtr devmode);

        [DllImport("gdi32")]
        public static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32")]
        public static extern int GetDeviceCaps(IntPtr hdc, int index);

        [DllImport("gdi32")]
        public static extern int StartDoc(IntPtr hdc, ref DOCINFO di);

        [DllImport("gdi32")]
        public static extern int EndDoc(IntPtr hdc);

        [DllImport("gdi32")]
        public static extern int StartPage(IntPtr hdc);

        [DllImport("gdi32")]
        public static extern int EndPage(IntPtr hdc);

        [DllImport("gdi32")]
        public static extern IntPtr CreateEnhMetaFile(string drv, string file, ref RECT frame, string desc);

        [DllImport("gdi32")]
        public static extern IntPtr CloseEnhMetaFile(IntPtr hdc);

        [DllImport("gdi32")]
        public static extern bool DeleteEnhMetaFile(IntPtr hemf);

        [DllImport("gdi32", CharSet = CharSet.Ansi, EntryPoint = "EnumEnhMetaFile")]
        public static extern bool EnumEnhMetaFile(IntPtr hdc, IntPtr hemf,
            EnumProc proc, IntPtr param, ref RECT frame);

        [DllImport("gdi32")]
        public static extern IntPtr CreatePalette(ref LOGPALETTE pal);

        [DllImport("gdi32")]
        public static extern IntPtr SelectPalette(IntPtr hdc, IntPtr pal, bool force);

        [DllImport("gdi32")]
        public static extern uint RealizePalette(IntPtr hdc);

        public delegate int EnumProc(IntPtr hdc, IntPtr table, IntPtr record, int size, IntPtr param);

        [StructLayout(LayoutKind.Sequential)]
        public struct DOCINFO { public int cbSize; public string docName; public string output; public string datatype; public int fwType; }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int left, top, right, bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct LOGPALETTE { public ushort ver; public ushort num; public uint entry; }

        public const int HORZRES = 8, VERTRES = 10, HORZSIZE = 4, VERTSIZE = 6;
        public const int LOGPIXELSX = 88, LOGPIXELSY = 90;

        public static byte[] ReadEmf(IntPtr hemf, int pageW, int pageH)
        {
            var frame = new RECT { left = 0, top = 0, right = pageW, bottom = pageH };
            var sink = new MemoryStream();
            EnumProc deliver = (hdc, table, record, size, param) =>
            {
                var buf = new byte[8];
                Marshal.Copy(record, buf, 0, 8);
                sink.Write(buf, 0, 8);
                var body = new byte[size - 8];
                Marshal.Copy(new IntPtr(record.ToInt32() + 8), body, 0, body.Length);
                sink.Write(body, 0, body.Length);
                return 1;
            };
            EnumEnhMetaFile(IntPtr.Zero, hemf, deliver, IntPtr.Zero, ref frame);
            return sink.ToArray();
        }
    }

    internal static class SPL
    {
        [DllImport("winspool.drv", CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern bool OpenPrinter(string name, out IntPtr handle, IntPtr defaults);

        [DllImport("winspool.drv", SetLastError = true)]
        public static extern bool ClosePrinter(IntPtr handle);

        [DllImport("winspool.drv", EntryPoint = "GetPrinterA", SetLastError = true)]
        public static extern bool GetPrinter(IntPtr handle, uint level, IntPtr data, uint size, out uint needed);

        public static bool CanOpen(string printer)
        {
            if (string.IsNullOrEmpty(printer)) return false;
            IntPtr h;
            if (!OpenPrinter(printer, out h, IntPtr.Zero)) return false;
            ClosePrinter(h);
            return true;
        }

        public static byte[] ReadDevMode(string printer)
        {
            IntPtr h;
            if (!OpenPrinter(printer, out h, IntPtr.Zero)) return null;
            try
            {
                uint need;
                GetPrinter(h, 2, IntPtr.Zero, 0, out need);
                if (need == 0) return null;
                IntPtr buf = Marshal.AllocHGlobal((int)need);
                try
                {
                    if (!GetPrinter(h, 2, buf, need, out need)) return null;
                    IntPtr pDev = Marshal.ReadIntPtr(buf, 0x40);
                    if (pDev == IntPtr.Zero) return null;
                    int dmSize = Marshal.ReadInt16(pDev, 0x24);
                    int extra = Marshal.ReadInt16(pDev, 0x26);
                    var mode = new byte[dmSize + extra];
                    Marshal.Copy(pDev, mode, 0, mode.Length);
                    return mode;
                }
                finally { Marshal.FreeHGlobal(buf); }
            }
            finally { ClosePrinter(h); }
        }

        public static void ReadNames(string printer, out string driver, out string port)
        {
            driver = port = null;
            IntPtr h;
            if (!OpenPrinter(printer, out h, IntPtr.Zero)) return;
            try
            {
                uint need;
                GetPrinter(h, 2, IntPtr.Zero, 0, out need);
                if (need == 0) return;
                IntPtr buf = Marshal.AllocHGlobal((int)need);
                try
                {
                    if (!GetPrinter(h, 2, buf, need, out need)) return;
                    IntPtr pDrv = Marshal.ReadIntPtr(buf, 0x44);
                    IntPtr pPort = Marshal.ReadIntPtr(buf, 0x48);
                    if (pDrv != IntPtr.Zero) driver = Marshal.PtrToStringUni(pDrv);
                    if (pPort != IntPtr.Zero) port = Marshal.PtrToStringUni(pPort);
                }
                finally { Marshal.FreeHGlobal(buf); }
            }
            finally { ClosePrinter(h); }
        }
    }

    #endregion

    #region Layout

    internal static class PI
    {
        public const int Size = 0x3C;
        public const int PD = 0x00, DockObject = 0x04, Preview = 0x08, Direct = 0x0C;
        public const int Continue = 0x10, CurPage = 0x14, NumPreview = 0x18;
        public const int StrPageDesc = 0x1C, UserData = 0x20;
        public const int RectL = 0x24, RectT = 0x28, RectR = 0x2C, RectB = 0x30;
        public const int OffsetPage = 0x34, Flags = 0x38;
    }

    internal static class PD
    {
        public const int DialogSize = 0x60;
        public const int DevMode = DialogSize + 0x08;
        public const int DevNames = DialogSize + 0x0C;
        public const int Hdc = DialogSize + 0x10;
        public const int Flags = DialogSize + 0x14;
        public const int FromPage = DialogSize + 0x18;
        public const int ToPage = DialogSize + 0x1A;
        public const int Copies = DialogSize + 0x1C;
    }

    internal static class DEVNAMES
    {
        public const int Driver = 0x04, Device = 0x06, Output = 0x08, TextStart = 0x0A;
    }

    internal static class VT
    {
        public const int IsSelected = 0x3F * 4;
        public const int PrepareDC = 0x4D * 4;
        public const int PreparePrinting = 0x4F * 4;
        public const int BeginPrinting = 0x50 * 4;
        public const int Print = 0x51 * 4;
        public const int EndPrinting = 0x52 * 4;
    }

    internal static class ViewOff
    {
        public const int Doc = 0x40, Gdi1 = 0x2B4, Gdi2 = 0x37C;
        public const int DevMode = 0x570, DevNames = 0x574, CacheMode = 0x578, CacheNames = 0x57C;
        public const int ColumnBase = 0x344, RowCount = 0x158, ColCount = 0x159;
        public const int DevModeHold = 0x158, DevNamesHold = 0x15C;
    }

    internal static class ProfileOff { public const int Base = 0x354; }
    internal static class Prop { public const int Printer = 23; }

    #endregion

    #region Delegates

    [UnmanagedFunctionPointer(CallingConvention.FastCall)]
    internal delegate int dPreparePrinting(IntPtr view, IntPtr pInfo);

    [UnmanagedFunctionPointer(CallingConvention.FastCall)]
    internal delegate void dPrepareDC(IntPtr view, IntPtr cdc, IntPtr pInfo);

    [UnmanagedFunctionPointer(CallingConvention.FastCall)]
    internal delegate void dPrint(IntPtr view, IntPtr cdc, IntPtr pInfo);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    internal delegate int dDeviceDefaults(IntPtr app, IntPtr pd);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    internal delegate int dGetCaps(IntPtr hdc, int index);

    #endregion

    /// <summary>
    /// Перехватывает печать CSheetView в 1С 7.7 и отдаёт PDF.
    /// </summary>
    public sealed class HeadlessPrintWrapper : IDisposable
    {
        const uint MEM_COMMIT_RESERVE = 0x3000, PAGE_EXECUTE_READWRITE = 0x40;
        const uint GMEM_MOVEABLE = 0x42;

        readonly IntPtr _view;
        readonly string _dumpDir;
        readonly string _pdfPath;
        readonly int _maxPages;

        IntPtr _vtable;
        IntPtr _pInfo, _pPD, _refDC, _pCdc, _gdi;
        byte[] _savedPrepare, _savedPrint;
        IntPtr _trampPrepare, _trampPrint;

        IntPtr _holdMode, _holdNames;
        string _printer, _driver, _port;
        string _savedProfilePrinter;

        readonly IPdfWriter _writer;
        readonly List<byte[]> _pages = new List<byte[]>();

        bool _armed, _disposed;

        public string ResolvedPrinter => _printer;
        public string ResolvedDriver => _driver;
        public string ResolvedPort => _port;
        public int PagesCaptured => _pages.Count;

        public HeadlessPrintWrapper(IntPtr sheetView, string dumpDir, string pdfPath, IPdfWriter writer, int maxPages = 2000)
        {
            if (sheetView == IntPtr.Zero) throw new ArgumentNullException(nameof(sheetView));
            _view = sheetView;
            _dumpDir = dumpDir;
            _pdfPath = pdfPath;
            _writer = writer ?? throw new ArgumentNullException(nameof(writer));
            _maxPages = maxPages;
        }

        // =================================================================
        // Точка входа
        // =================================================================

        public bool Print()
        {
            try
            {
                _pages.Clear();
                _vtable = Marshal.ReadIntPtr(_view, 0);
                _gdi = new IntPtr(_view.ToInt32() + ViewOff.Gdi2);

                _holdMode = LoadResource(Path.Combine(_dumpDir, "devmode.bin"));
                _holdNames = LoadResource(Path.Combine(_dumpDir, "devnames.bin"));
                if (_holdMode == IntPtr.Zero || _holdNames == IntPtr.Zero)
                    return Fail("нет devmode.bin/devnames.bin в " + _dumpDir);

                ReadIni(Path.Combine(_dumpDir, "printer.ini"));
                ResolveLive();

                _pInfo = AllocPrintInfo();
                _pPD = Marshal.ReadIntPtr(_pInfo, PI.PD);
                if (_pPD == IntPtr.Zero) return Fail("конструктор не создал CPrintDialog");

                _savedProfilePrinter = ReadProfileString(Prop.Printer);
                WriteProfileString(Prop.Printer, _printer);

                PinDeviceBlocks();
                InstallHooks();

                Marshal.WriteInt32(_pInfo, PI.Direct, 1);            // диалог печати не открывать
                Marshal.WriteInt32(_pInfo, PI.Preview, 0);

                var prepare = Bind<dPreparePrinting>(VT.PreparePrinting);
                if (prepare(_view, _pInfo) == 0) return Fail("OnPreparePrinting вернул 0");

                if (!TakeReferenceDc()) return Fail("нет принтерного DC");
                FillRectDraw();

                _pCdc = MakeFakeCdc();
                var begin = Bind<dPrepareDC>(VT.BeginPrinting);
                begin(_view, _pCdc, _pInfo);

                int from = Math.Max((short)0x01, Marshal.ReadInt16(_pPD, PD.FromPage));
                int to = Math.Max(from, Marshal.ReadInt16(_pPD, PD.ToPage));
                CapturePages(from, to);

                var end = Bind<dPrepareDC>(VT.EndPrinting);
                end(_view, _pCdc, _pInfo);

                if (_pages.Count == 0) return Fail("не отрисовано ни одной страницы");
                _writer.Write(_pdfPath, _pages, _pageW, _pageH);
                return true;
            }
            catch (Exception e) { return Fail(e.ToString()); }
            finally
            {
                try { UninstallHooks(); } catch { }
                if (_pInfo != IntPtr.Zero) { FreePrintInfo(_pInfo); _pInfo = IntPtr.Zero; }
            }
        }

        int _pageW, _pageH;

        void CapturePages(int from, int to)
        {
            var prep = Bind<dPrepareDC>(VT.PrepareDC);
            var print = Bind<dPrint>(VT.Print);
            byte[] page = null;
            int drawn = 0;

            for (int p = from; p <= to && drawn < _maxPages; p++, drawn++)
            {
                if (Marshal.ReadInt32(_pInfo, PI.Continue) == 0 && p > from) break;
                Marshal.WriteInt32(_pInfo, PI.CurPage, p);

                IntPtr emfDc = StartPageEmf();
                SwapHdc(_pCdc, emfDc);
                prep(_view, _pCdc, _pInfo);
                print(_view, _pCdc, _pInfo);
                page = EndPageEmf(emfDc);

                if (page != null && page.Length > 0) _pages.Add(page);
                if (Marshal.ReadInt32(_pInfo, PI.Continue) == 0) break;
            }
        }

        // =================================================================
        // Прямой режим: устройство
        // =================================================================

        void ResolveLive()
        {
            if (_printer != null && SPL.CanOpen(_printer))
            {
                byte[] live = SPL.ReadDevMode(_printer);
                if (live != null) FreeResource(ref _holdMode);
                if (live != null) _holdMode = MakeGlb(live);
                SPL.ReadNames(_printer, out _driver, out _port);
                return;
            }

            // живого принтера нет: имя и параметры берём из дампа
            if (_printer == null && _driver == null)
            {
                // в printer.ini уже прочитано
            }
        }

        int DeviceDefaultsHook(IntPtr app, IntPtr pd)
        {
            if (!_armed) return _origDeviceDefaults(app, pd);
            try
            {
                Marshal.WriteIntPtr(pd, PD.DevMode, _holdMode);
                Marshal.WriteIntPtr(pd, PD.DevNames, _holdNames);
                return 1;
            }
            catch { return 0; }
        }

        dDeviceDefaults _origDeviceDefaults;

        void PinDeviceBlocks()
        {
            // Одиночная фиксация: блоки живут весь процесс, их нельзя отдать в оба слота.
            if (_holdMode == IntPtr.Zero) _holdMode = MakeGlb(new byte[0xDC]);
            if (_holdNames == IntPtr.Zero) _holdNames = MakeGlb(BuildDevNames(_driver, _printer, _port));
        }

        bool TakeReferenceDc()
        {
            IntPtr hMode = Marshal.ReadIntPtr(_pPD, PD.DevMode);
            IntPtr hNames = Marshal.ReadIntPtr(_pPD, PD.DevNames);
            if (hMode == IntPtr.Zero || hNames == IntPtr.Zero) return false;
            _refDC = Marshal.ReadIntPtr(_pPD, PD.Hdc);
            if (_refDC == IntPtr.Zero)
            {
                IntPtr devMode = K.GlobalLock(hMode);
                try { _refDC = GDI.CreateDC(_driver, _printer, _port, devMode); }
                finally { K.GlobalUnlock(hMode); }
            }
            if (_refDC == IntPtr.Zero) return false;
            _pageW = GDI.GetDeviceCaps(_refDC, GDI.HORZRES);
            _pageH = GDI.GetDeviceCaps(_refDC, GDI.VERTRES);
            return _pageW > 0 && _pageH > 0;
        }

        void FillRectDraw()
        {
            Marshal.WriteInt32(_pInfo, PI.RectL, 0);
            Marshal.WriteInt32(_pInfo, PI.RectT, 0);
            Marshal.WriteInt32(_pInfo, PI.RectR, _pageW);
            Marshal.WriteInt32(_pInfo, PI.RectB, _pageH);
        }

        // =================================================================
        // Патч vtable
        // =================================================================

        void InstallHooks()
        {
            PatchSlot(VT.Print, Marshal.GetFunctionPointerForDelegate((dPrintDetour)OnPrintDetour), ref _savedPrint, ref _trampPrint);
            _armed = true;
        }

        void UninstallHooks()
        {
            _armed = false;
            if (_trampPrint != IntPtr.Zero)
            {
                WriteSlot(VT.Print, _trampPrint);
                _trampPrint = IntPtr.Zero;
            }
        }

        void PatchSlot(int slot, IntPtr detour, ref byte[] saved, ref IntPtr tramp)
        {
            IntPtr addr = Marshal.ReadIntPtr(_vtable, slot);
            saved = new byte[16];
            Marshal.Copy(addr, saved, 0, saved.Length);

            tramp = K.VirtualAlloc(IntPtr.Zero, new IntPtr(32), MEM_COMMIT_RESERVE, PAGE_EXECUTE_READWRITE);
            Marshal.Copy(saved, 0, tramp, saved.Length);
            IntPtr back = new IntPtr(tramp.ToInt32() + saved.Length);
            WriteJump(back, new IntPtr(addr.ToInt32() + saved.Length));

            WriteSlot(slot, detour);
        }

        void WriteSlot(int slot, IntPtr fn) => Marshal.WriteIntPtr(_vtable, slot, fn);

        static void WriteJump(IntPtr at, IntPtr to)
        {
            uint old;
            K.VirtualProtect(at, new IntPtr(5), PAGE_EXECUTE_READWRITE, out old);
            Marshal.WriteByte(at, 0, 0xE9);
            Marshal.WriteInt32(at, 1, to.ToInt32() - at.ToInt32() - 5);
        }

        [UnmanagedFunctionPointer(CallingConvention.FastCall)]
        delegate void dPrintDetour(IntPtr view, IntPtr cdc, IntPtr pInfo);

        void OnPrintDetour(IntPtr view, IntPtr cdc, IntPtr pInfo)
        {
            // Пока цикл сами ведём мы, страницу рисуем прямо здесь.
            DrawDetour(view, cdc, pInfo);
        }

        void DrawDetour(IntPtr view, IntPtr cdc, IntPtr pInfo)
        {
            var real = Marshal.GetDelegateForFunctionPointer<dPrint>(_trampPrint);
            real(view, cdc, pInfo);
        }

        T Bind<T>(int slot) where T : class
        {
            IntPtr fn = Marshal.ReadIntPtr(_vtable, slot);
            return (T)(object)Marshal.GetDelegateForFunctionPointer(fn, typeof(T));
        }

        // =================================================================
        // CDC
        // =================================================================

        IntPtr MakeFakeCdc()
        {
            IntPtr cdc = Marshal.ReadIntPtr(_gdi, 4);       // CSheetGDI.m_CDC по +4
            if (cdc == IntPtr.Zero) throw new InvalidOperationException("CSheetGDI.m_CDC пуст");
            return cdc;
        }

        static void SwapHdc(IntPtr cdc, IntPtr hdc)
        {
            Marshal.WriteIntPtr(cdc, 0, hdc);              // CDC.m_hDC
            Marshal.WriteIntPtr(cdc, 4, hdc);              // CDC.m_hAttribDC
        }

        IntPtr StartPageEmf()
        {
            var frame = new GDI.RECT { left = 0, top = 0, right = _pageW, bottom = _pageH };
            var di = new GDI.DOCINFO { cbSize = Marshal.SizeOf(typeof(GDI.DOCINFO)), docName = "1C" };
            IntPtr dc = GDI.CreateEnhMetaFile(_driver, null, ref frame, "1C headless");
            GDI.StartDoc(dc, ref di);
            GDI.StartPage(dc);
            return dc;
        }

        byte[] EndPageEmf(IntPtr dc)
        {
            GDI.EndPage(dc);
            GDI.EndDoc(dc);
            IntPtr hemf = GDI.CloseEnhMetaFile(dc);
            if (hemf == IntPtr.Zero) return null;
            try { return GDI.ReadEmf(hemf, _pageW, _pageH); }
            finally { GDI.DeleteEnhMetaFile(hemf); }
        }

        // =================================================================
        // CPrintInfo
        // =================================================================

        static IntPtr _mfc;

        static IntPtr Mfc()
        {
            if (_mfc == IntPtr.Zero)
            {
                _mfc = K.GetModuleHandleA("mfc42.dll");
                if (_mfc == IntPtr.Zero) throw new DllNotFoundException("mfc42.dll");
            }
            return _mfc;
        }

        IntPtr AllocPrintInfo()
        {
            IntPtr p = Marshal.AllocHGlobal(PI.Size);
            for (int i = 0; i < PI.Size; i += 4) Marshal.WriteInt32(p, i, 0);
            IntPtr fn = K.GetProcAddress(Mfc(), "??0CPrintInfo@@QAE@XZ");
            var ctor = Marshal.GetFunctionPointerForDelegate(
                new dThisPtr((IntPtr self) => self));
            // thiscall: this в ecx
            var call = BindCtor(fn);
            return call(p);
        }

        delegate IntPtr dThisPtr(IntPtr self);

        static dThisPtr BindCtor(IntPtr fn)
        {
            // __thiscall одним аргументом эквивалентен fastcall(ecx)
            return (dThisPtr)Marshal.GetDelegateForFunctionPointer(fn, typeof(CtorStub));
        }

        [UnmanagedFunctionPointer(CallingConvention.FastCall)]
        delegate IntPtr CtorStub(IntPtr self);

        void FreePrintInfo(IntPtr p)
        {
            IntPtr fn = K.GetProcAddress(Mfc(), "??1CPrintInfo@@QAE@XZ");
            var dtor = Marshal.GetDelegateForFunctionPointer<CtorStub>(fn);
            dtor(p);
            Marshal.FreeHGlobal(p);
        }

        // =================================================================
        // Профиль
        // =================================================================

        IntPtr Profile()
        {
            IntPtr doc = Marshal.ReadIntPtr(_view, ViewOff.Doc);
            return new IntPtr(doc.ToInt32() + ProfileOff.Base);
        }

        string ReadProfileString(int key)
        {
            IntPtr fn = K.GetProcAddress(LoadFrame(), "?GetStringProp@CProfile7@@QBEABVCString@@H@Z");
            if (fn == IntPtr.Zero) return null;
            var get = Marshal.GetDelegateForFunctionPointer<GetStrStub>(fn);
            IntPtr cstr = Marshal.AllocHGlobal(8);
            try
            {
                get(cstr, Profile(), key);              // скрытый ret-ptr первым
                IntPtr data = Marshal.ReadIntPtr(cstr, 0);
                return data == IntPtr.Zero ? null : Marshal.PtrToStringAnsi(data);
            }
            finally { Marshal.FreeHGlobal(cstr); }
        }

        void WriteProfileString(int key, string value)
        {
            IntPtr fn = K.GetProcAddress(LoadFrame(), "?SetStringProp@CProfile7@@QAEXHPBD@Z");
            if (fn == IntPtr.Zero) return;
            var set = Marshal.GetDelegateForFunctionPointer<SetStrStub>(fn);
            set(Profile(), key, value);
        }

        [UnmanagedFunctionPointer(CallingConvention.FastCall)]
        delegate void GetStrStub(IntPtr ret, IntPtr self, int key);

        [UnmanagedFunctionPointer(CallingConvention.FastCall)]
        delegate void SetStrStub(IntPtr self, int key, string value);

        static IntPtr _frame;
        static IntPtr LoadFrame()
        {
            if (_frame == IntPtr.Zero) _frame = K.GetModuleHandleA("frame.dll");
            return _frame;
        }

        // =================================================================
        // Вспомогательное
        // =================================================================

        static IntPtr MakeGlb(byte[] data)
        {
            IntPtr h = K.GlobalAlloc(GMEM_MOVEABLE, new IntPtr(data.Length + 8));
            IntPtr p = K.GlobalLock(h);
            Marshal.Copy(data, 0, p, data.Length);
            K.GlobalUnlock(h);
            return h;
        }

        static IntPtr LoadResource(string path)
        {
            if (!File.Exists(path)) return IntPtr.Zero;
            return MakeGlb(File.ReadAllBytes(path));
        }

        static void FreeResource(ref IntPtr h)
        {
            if (h == IntPtr.Zero) return;
            int locked = K.GlobalFlags(h) & 0xFF;
            for (int i = 0; i < locked; i++) K.GlobalUnlock(h);
            K.GlobalFree(h);
            h = IntPtr.Zero;
        }

        static byte[] BuildDevNames(string driver, string device, string port)
        {
            driver = driver ?? ""; device = device ?? ""; port = port ?? "";
            byte[] d = Encoding.Default.GetBytes(driver), n = Encoding.Default.GetBytes(device), p = Encoding.Default.GetBytes(port);
            int total = 0x0A + d.Length + 1 + n.Length + 1 + p.Length + 1;
            var b = new byte[total];
            BitConverter.GetBytes((short)0x0A).CopyTo(b, 0);
            BitConverter.GetBytes((short)(0x0A)).CopyTo(b, 0x02);
            BitConverter.GetBytes((short)(0x0A + d.Length + 1)).CopyTo(b, 0x04);
            BitConverter.GetBytes((short)(0x0A + d.Length + 1 + n.Length + 1)).CopyTo(b, 0x06);
            int o = 0x0A;
            Array.Copy(d, 0, b, o, d.Length); o += d.Length + 1;
            Array.Copy(n, 0, b, o, n.Length); o += n.Length + 1;
            Array.Copy(p, 0, b, o, p.Length);
            return b;
        }

        void ReadIni(string path)
        {
            if (!File.Exists(path)) return;
            foreach (string line in File.ReadAllLines(path))
            {
                int eq = line.IndexOf('=');
                if (eq < 0) continue;
                string k = line.Substring(0, eq).Trim(), v = line.Substring(eq + 1).Trim();
                if (k == "Name") _printer = v;
                else if (k == "Driver") _driver = v;
                else if (k == "Port") _port = v;
            }
        }

        const int DM_ORIENTATION = 1, DM_PAPERSIZE = 2, DM_PAPERLENGTH = 4, DM_PAPERWIDTH = 8,
          DM_SCALE = 0x10, DM_COPIES = 0x100, DM_DEFAULTSOURCE = 0x200,
          DM_PRINTQUALITY = 0x400, DM_COLOR = 0x800, DM_DUPLEX = 0x1000,
          DM_YRESOLUTION = 0x2000, DM_TTOPTION = 0x4000, DM_COLLATE = 0x8000,
          DM_FORMNAME = 0x10000;

        static byte[] BuildMinimalDevModeA(string printer, short orient, short paper,
                                  short widthMm, short lengthMm, short dpi, short copies)
        {
            const int dmSize = 0xDC;
            var b = new byte[dmSize];

            var nm = Encoding.Default.GetBytes(printer);
            Array.Copy(nm, 0, b, 0x00, Math.Min(nm.Length, 31));   // dmDeviceName, ANSI

            Write16(b, 0x20, 0x0401);        // dmSpecVersion, 4.x
            Write16(b, 0x24, dmSize);        // dmSize
            Write16(b, 0x26, 0);             // dmDriverExtra
            Write32(b, 0x28,
                DM_ORIENTATION | DM_PAPERSIZE | DM_PAPERWIDTH | DM_PAPERLENGTH |
                DM_PRINTQUALITY | DM_COPIES | DM_COLOR | DM_DUPLEX);   // dmFields

            Write16(b, 0x2E, orient);        // dmOrientation, 1 portrait 2 landscape
            Write16(b, 0x30, paper);         // dmPaperSize, 9 = A4, 0 = задаётся вручную
            Write16(b, 0x32, lengthMm);      // dmPaperLength, 0.1 мм
            Write16(b, 0x34, widthMm);       // dmPaperWidth,  0.1 мм
            Write16(b, 0x36, 100);           // dmScale, проценты
            Write16(b, 0x38, copies);        // dmCopies
            Write16(b, 0x3A, 7);             // dmDefaultSource, 7 = auto
            Write16(b, 0x3C, dpi);           // dmPrintQuality, при DM_RESERVED1 это DPI
            Write16(b, 0x3E, 1);             // dmColor, 1 = монохром
            Write16(b, 0x40, 1);             // dmDuplex, 1 = односторонняя
            Write16(b, 0x42, dpi);           // dmYResolution
            Write16(b, 0x44, 0);             // dmTTOption, 0 = растром
            Write16(b, 0x46, 0);             // dmCollate, 0 = не разбирать
            return b;

            static void Write16(byte[] a, int o, int v) => BitConverter.GetBytes((short)v).CopyTo(a, o);
            static void Write32(byte[] a, int o, int v) => BitConverter.GetBytes(v).CopyTo(a, o);
        }

        static bool Fail(string why) { S.Log("[HeadlessPrint] " + why); return false; }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { UninstallHooks(); } catch { }
            if (_pInfo != IntPtr.Zero) { FreePrintInfo(_pInfo); _pInfo = IntPtr.Zero; }
            if (_writer is IDisposable d) d.Dispose();
            GC.SuppressFinalize(this);
        }
    }

    public interface IPdfWriter : IDisposable
    {
        void Write(string path, IList<byte[]> emfPages, int width, int height);
    }

    internal static class S
    {
        public static void Log(string msg) =>
            System.Diagnostics.Debug.WriteLine(msg);
    }
}