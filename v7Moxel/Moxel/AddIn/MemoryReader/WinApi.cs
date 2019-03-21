using System;
using System.Runtime.InteropServices;

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

        public static uint MakeIntResource(uint ResID)
        {
            return (ResID >> 4) + 1;
        }
    }
}
