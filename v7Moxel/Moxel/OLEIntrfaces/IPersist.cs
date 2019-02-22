using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("0000010C-0000-0000-C000-000000000046")]
    public interface IPersist
    {
        int GetClassID([Out] out Guid pClassID);
    }
}
