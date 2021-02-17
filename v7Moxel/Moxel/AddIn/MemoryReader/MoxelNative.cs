using System;
using System.Runtime.InteropServices;

namespace Moxel
{
    public static class MoxelNative
    {

        [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate void SerializeDelegate(IntPtr pObj, IntPtr pArch);

        private static readonly IntPtr HMoxel = WinApi.GetModuleHandle("moxel.dll");

        public static SerializeDelegate GetSerializer(string entryPoint)
        {
            return GetDelegate<SerializeDelegate>(entryPoint);
        }

        public static T GetDelegate<T>(string funcName)
        {
            return Marshal.GetDelegateForFunctionPointer<T>(WinApi.GetProcAddress(HMoxel, funcName));
        }
    }
}
