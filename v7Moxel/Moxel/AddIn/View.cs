using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;
using System.Windows.Data;

static class View
{
    [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
    public delegate int dOnPreparePrinting(IntPtr view, IntPtr pInfo);

    [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
    public delegate void dOnPrint(IntPtr view, IntPtr pDC, IntPtr pInfo);

    [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
    public delegate void dOnPrepareDC(IntPtr view, IntPtr pDC, IntPtr pInfo);

    [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
    public delegate void dOnBeginPrinting(IntPtr view, IntPtr pDC, IntPtr pInfo);

    [UnmanagedFunctionPointer(CallingConvention.ThisCall)]
    public delegate void dOnEndPrinting(IntPtr view, IntPtr pDC, IntPtr pInfo);

    public static readonly int VT_OnPrepareDC = 0x4D * 4;
    public static readonly int VT_OnPreparePrinting = 0x4F * 4;
    public static readonly int VT_OnBeginPrinting = 0x50 * 4;
    public static readonly int VT_OnPrint = 0x51 * 4;
    public static readonly int VT_OnEndPrinting = 0x52 * 4;


    public static T Vt<T>(IntPtr obj, int slot)
    {
        IntPtr vt = Marshal.ReadIntPtr(obj, 0);
        IntPtr fn = Marshal.ReadIntPtr(vt, slot);
        return (T)(object)Marshal.GetDelegateForFunctionPointer(fn, typeof(T));
    }

    public const int GetDocPtr = 0x40;    // CSheetDoc* по view+0x40
    public const int ProfileOff = 0x354;   // CProfile7 внутри CSheetDoc
    public const int DevModeOff = 0x570;   // HGLOBAL m_hDevMode
    public const int DevNamesOff = 0x574;   // HGLOBAL m_hDevNames
    public const int Gdi1Off = 0x2B4;   // CSheetGDI, первый
    public const int Gdi2Off = 0x37C;   // CSheetGDI, второй, тот что печатает

    public static IntPtr Doc(IntPtr v) => Marshal.ReadIntPtr(v, GetDocPtr);
    public static IntPtr Profile(IntPtr v) => new IntPtr(Doc(v).ToInt32() + ProfileOff);
}

static class PDA
{
    public const int OffDevMode = 0x08;
    public const int OffDevNames = 0x0C;
}

static class PrintDlg
{
    public const int Size = 0x60;
    public const int OffFlags = Size + 0x14;
    public const int OffHdc = Size + 0x10;
    public const int OffFrom = Size + 0x18;
    public const int OffTo = Size + 0x1A;
    public const int OffCopies = Size + 0x1C;
}

static class PrintInfo
{
    public const int OffPPD = 0x00;
    public const int OffCurPage = 0x0C;
    public const int OffRect = 0x10;
    public const int OffPreview = 0x20;
    public const int OffMaxPage = 0x0C;   // внутри pPD совпадает, но не путать
    public const int OffFromPage = 0x10;   // m_nFromPage, если понадобится
    public const int OffToPage = 0x12;
}