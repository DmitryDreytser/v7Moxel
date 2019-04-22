using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Moxel
{
    public static class WinApi
    {
        [DllImport("Kernel32", EntryPoint = "GetModuleHandleW", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string dllname);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr GetProcAddress(IntPtr HModule, [MarshalAs(UnmanagedType.LPStr), In] string funcName);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr GetProcAddress(IntPtr HModule, int ordinal);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool VirtualProtectEx(IntPtr hProcess, IntPtr lpAddress, IntPtr dwSize, uint flNewProtect, out uint lpflOldProtect);

        [DllImport("kernel32.dll")]
        public static extern bool FlushInstructionCache(IntPtr hProcess, IntPtr lpBaseAddress, IntPtr dwSize);

        [DllImport("kernel32.dll")]
        public static extern IntPtr InterlockedExchange(IntPtr Target, IntPtr Value);

        [DllImport("kernel32.dll",EntryPoint = "FindResourceA")]
        public static extern IntPtr FindResource(IntPtr hModule, uint nID, int Type);

        [DllImport("kernel32.dll")]
        public static extern IntPtr LoadResource(IntPtr hModule, IntPtr HRSRC);

        [DllImport("kernel32.dll")]
        public static extern int SizeofResource(IntPtr hModule, IntPtr HRSRC);

        [DllImport("kernel32.dll")]
        public static extern IntPtr LockResource(IntPtr HGlobal);


        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("User32.dll")]
        public static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, StringBuilder lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, [MarshalAs(UnmanagedType.LPStr)] string lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, ref IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);


        [DllImport("gdi32.dll")]
        public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        const int LOGPIXELSX = 88;
        const int LOGPIXELSY = 90;
        const float UNITS_PER_INCH = 1440f;

        public static float GetUnitsPerPixel()
        {
            IntPtr hDc = GetDC(IntPtr.Zero);
            float m_UnitsPerPixel = UNITS_PER_INCH / GetDeviceCaps(hDc, LOGPIXELSX);
            ReleaseDC(IntPtr.Zero, hDc);
            return m_UnitsPerPixel;

        }

        public static uint MakeIntResource(uint ResID)
        {
            return (ResID >> 4) + 1;
        }

        public static T GetDelegate<T>(string ModuleName, string FuncName)
        {
            return Marshal.GetDelegateForFunctionPointer<T>(GetProcAddress(GetModuleHandle(ModuleName), FuncName));
        }

        public static T GetDelegate<T>(string ModuleName, int FuncOrdinal)
        {
            return Marshal.GetDelegateForFunctionPointer<T>(GetProcAddress(GetModuleHandle(ModuleName), FuncOrdinal));
        }
    }
}
