using System;
using System.Runtime.InteropServices;

namespace Moxel
{
    public static class MoxelNative
    {

        [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate void SerializeDelegate(IntPtr pObj, IntPtr pArch);

        static IntPtr hMoxel = WinApi.GetModuleHandle("moxel.dll");

        public static SerializeDelegate GetSerializer(string EntryPoint)
        {
            return GetDelegate<SerializeDelegate>(EntryPoint);
        }

        public static T GetDelegate<T>(string FuncName)
        {
            return Marshal.GetDelegateForFunctionPointer<T>(WinApi.GetProcAddress(hMoxel, FuncName));
        }
    }
}
