using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [ComImport]
    [Guid("55272A00-42CB-11CE-8135-00AA004BB851")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyBag
    {
        [PreserveSig]
        int Read(
          [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
          [Out, MarshalAs(UnmanagedType.Struct)] out object pVar,
          [In, MarshalAs(UnmanagedType.Interface)] IErrorLog pErrorLog
        );

        [PreserveSig]
        int Write(
          [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
          [In, MarshalAs(UnmanagedType.Struct)] ref object pVar
        );
    }
}
