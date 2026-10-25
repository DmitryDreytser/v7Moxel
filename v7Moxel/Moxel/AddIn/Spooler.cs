using System;
using System.Runtime.InteropServices;
using System.Text;

struct tagPDA { public IntPtr hDevMode; public IntPtr hDevNames; }

static class Spooler
{
    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool OpenPrinterW(string name, out IntPtr h, IntPtr pd);

    [DllImport("winspool.drv", SetLastError = true)]
    static extern bool ClosePrinter(IntPtr h);

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool GetPrinterW(IntPtr h, int lvl, IntPtr buf, int cb, out int need);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GlobalAlloc(uint flags, IntPtr bytes);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GlobalFree(IntPtr h);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GlobalLock(IntPtr h);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GlobalUnlock(IntPtr h);

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool GetDefaultPrinterW(System.Text.StringBuilder name, ref int cch); // cch = 0 даёт размер

    public static string GetDefaultPrinter()
    {
        var name = new StringBuilder();
        int cch = 0;
        if(GetDefaultPrinterW(name, ref cch))
            return name.ToString();
        return string.Empty;
    }

    // PRINTER_INFO_2, только нужные поля для x86
    const int PI2_pDevMode = 0x40;
    const int PI2_pDriverName = 0x44;
    const int PI2_pPortName = 0x48;

    public static string LoadPrinterName(string preferred)
    {
        if (!string.IsNullOrEmpty(preferred) && CanOpen(preferred)) return preferred;
        int n = 0;
        GetDefaultPrinterW(null, ref n);
        var sb = new System.Text.StringBuilder(n > 1 ? n : 260);
        int cap = sb.Capacity;
        if (n > 0 && GetDefaultPrinterW(sb, ref cap)) { return sb.ToString(); }
        return null;
    }

    static bool CanOpen(string name)
    {
        IntPtr h;
        if (!OpenPrinterW(name, out h, IntPtr.Zero)) return false;
        ClosePrinter(h);
        return true;
    }

    public static bool Build(IntPtr hDevMode, IntPtr hDevNames, string name, out string printerName)
    {
        // hDevMode и hDevNames приходят уже выделенными GlobalAlloc и заполненными
        return CaptureInfo(name, hDevMode, hDevNames, out printerName);
    }

    // Открывает принтер, читает уровень 2 и переносит pDevMode, pDriverName, pPortName в наши блоки
    static bool CaptureInfo(string name, IntPtr hOutMode, IntPtr hOutNames, out string printerName)
    {
        IntPtr hP;
        printerName = null;
        if (!OpenPrinterW(name, out hP, IntPtr.Zero)) return false;
        try
        {
            int need;
            GetPrinterW(hP, 2, IntPtr.Zero, 0, out need);
            if (need <= 0) return false;
            IntPtr buf = Marshal.AllocHGlobal(need);
            try
            {
                if (!GetPrinterW(hP, 2, buf, need, out need)) return false;
                IntPtr pDev = Marshal.ReadIntPtr(buf, PI2_pDevMode);
                IntPtr pDrv = Marshal.ReadIntPtr(buf, PI2_pDriverName);
                if (pDev == IntPtr.Zero || pDrv == IntPtr.Zero) return false;

                // DEVMODE: dmSize в начале, сразу за ним dmDriverExtra
                ushort dmSize = (ushort)Marshal.ReadInt16(pDev, 0x00);
                ushort extra = (ushort)Marshal.ReadInt16(pDev, 0x02);
                int modeLen = dmSize + extra;
                if (modeLen < 0x100 || modeLen > 0x10000) return false;

                IntPtr srcMod = GlobalLock(pDev);
                IntPtr dstMod = GlobalLock(hOutMode);
                if (srcMod == IntPtr.Zero || dstMod == IntPtr.Zero) return false;
                try { Copy(srcMod, dstMod, modeLen); }
                finally { GlobalUnlock(hOutMode); GlobalUnlock(pDev); }

                // DEVNAMES { wDriverOffset, wDeviceOffset, wOutputOffset, wDefault } с ANSI-строками
                IntPtr pPort = Marshal.ReadIntPtr(buf, PI2_pPortName);
                byte[] drvS = AnsiZ(pDrv);
                byte[] devS = AnsiZ(Marshal.StringToHGlobalAnsi(name));
                byte[] outS = pPort != IntPtr.Zero ? AnsiZ(pPort) : new byte[] { (byte)'F', (byte)'I', (byte)'L', (byte)'E', (byte)':', 0 };
                int oDriver = 8;
                int oDevice = oDriver + drvS.Length;
                int oOutput = oDevice + devS.Length;
                int len = oOutput + outS.Length + 2;

                IntPtr dstN = GlobalLock(hOutNames);
                if (dstN == IntPtr.Zero) return false;
                try
                {
                    for (int i = 0; i < len; i++) Marshal.WriteByte(dstN, i, 0);
                    Marshal.WriteInt16(dstN, 0, (short)oDriver);
                    Marshal.WriteInt16(dstN, 2, (short)oDevice);
                    Marshal.WriteInt16(dstN, 4, (short)oOutput);
                    Marshal.WriteInt16(dstN, 6, 0);
                    Marshal.Copy(drvS, 0, new IntPtr(dstN.ToInt32() + oDriver), drvS.Length);
                    Marshal.Copy(devS, 0, new IntPtr(dstN.ToInt32() + oDevice), devS.Length);
                    Marshal.Copy(outS, 0, new IntPtr(dstN.ToInt32() + oOutput), outS.Length);
                }
                finally { GlobalUnlock(hOutNames); }

                printerName = name;

                return true;
            }
            finally { Marshal.FreeHGlobal(buf); }
        }
        finally { ClosePrinter(hP); }
    }

    static void Copy(IntPtr src, IntPtr dst, int n)
    {
        for (int i = 0; i < n; i += 4)
        {
            int v = Marshal.ReadInt32(src, i);
            if (i + 4 <= n) Marshal.WriteInt32(dst, i, v);
        }
    }

    static byte[] AnsiZ(IntPtr p) => AnsiZ(Marshal.PtrToStringAnsi(p));
    static byte[] AnsiZ(string s)
    {
        var b = System.Text.Encoding.Default.GetBytes(s ?? "");
        var r = new byte[b.Length + 1];
        Array.Copy(b, r, b.Length);
        return r;
    }
}