using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;

namespace Ole
{
    public enum OLEUPDATE
    {
        OLEUPDATE_ALWAYS = 1,
        OLEUPDATE_ONCALL = 3
    };

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000011d-0000-0000-C000-000000000046")]
    public interface IOleLink
    {
        HRESULT  SetUpdateOptions([In]  int dwUpdateOpt);

        HRESULT  GetUpdateOptions([Out] out int  pdwUpdateOpt);

        HRESULT  SetSourceMoniker([In, MarshalAs(UnmanagedType.Interface)] IMoniker  pmk, [In] Guid rclsid);

        HRESULT  GetSourceMoniker([Out, MarshalAs(UnmanagedType.Interface)] out IMoniker  ppmk);

        HRESULT  SetSourceDisplayName([In, MarshalAs(UnmanagedType.LPWStr)] string pszStatusText);

        HRESULT  GetSourceDisplayName([Out, MarshalAs(UnmanagedType.LPWStr)] out string ppszDisplayName);
        HRESULT  BindToSource([In] int bindflags,[In, MarshalAs(UnmanagedType.Interface)] IBindCtx  pbc);
        HRESULT BindIfRunning([MarshalAs(UnmanagedType.Interface)] IOleLink pthis);
        HRESULT  GetBoundSource([Out, MarshalAs(UnmanagedType.IUnknown)] out IntPtr ppunk);
        HRESULT  UnbindSource();

        HRESULT  Update([In, MarshalAs(UnmanagedType.Interface)] IBindCtx  pbc);
    }
}
