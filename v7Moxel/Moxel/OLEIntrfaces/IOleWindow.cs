using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [ComImport, Guid("00000114-0000-0000-C000-000000000046")]
    [InterfaceType(1)]
    public interface IOleWindow
    {
        //
        // Parameters:
        //   fEnterMode:
        void ContextSensitiveHelp(int fEnterMode);
        //
        // Parameters:
        //   phwnd:
        void GetWindow(out IntPtr phwnd);
    }
}
