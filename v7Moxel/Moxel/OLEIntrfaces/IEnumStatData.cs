using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [Guid("00000105-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IEnumStatData
    {
        [PreserveSig]
        int Next(
            [In] uint celt,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] [Out] StatData[] rgelt,
            [MarshalAs(UnmanagedType.LPArray)] [Out] uint[] pceltFetched);
        [PreserveSig]
        int Skip(
            [In] uint celt);
        [PreserveSig]
        int Reset();
        [PreserveSig]
        int Clone(
            [MarshalAs(UnmanagedType.Interface)] out IEnumStatData ppEnum);
    }
}
