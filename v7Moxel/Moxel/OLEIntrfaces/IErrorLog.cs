using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [ComImport]
    [InterfaceType(1)]
    [Guid("3127CA40-446E-11CE-8135-00AA004BB851")]
    public interface IErrorLog
    {
        void AddError(string pszPropName, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo);
    }
}
