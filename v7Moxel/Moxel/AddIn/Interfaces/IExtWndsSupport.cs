using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AddIn
{
    [ComImport]
    [Guid("EFE19EA0-09E4-11D2-A601-008048DA00DE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IExtWndsSupport
    {
        void GetAppMainFrame(
            [Out] out IntPtr hwnd);

        void GetAppMDIFrame(
            [Out] out IntPtr hwnd);

        void CreateAddInWindow(
            [In] string bstrProgID,
            [In] string bstrWindowName,
            [In] int lStyles,
            [In] int lExStyles,
            [In, Out] ref System.Drawing.Size rctSize,
            [In] int Flags,
            [In, Out] ref IntPtr pHwnd,
            [In, Out, MarshalAs(UnmanagedType.IDispatch)] ref object pDisp);
    }
}
