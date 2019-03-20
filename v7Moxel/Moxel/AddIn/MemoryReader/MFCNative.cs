using System;
using System.Runtime.InteropServices;

namespace Moxel
{
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
}
