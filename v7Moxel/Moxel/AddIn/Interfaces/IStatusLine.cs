using System;
using System.Runtime.InteropServices;

namespace AddIn
{
    [ComVisible(true)]
    [Guid("AB634005-F13D-11D0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStatusLine
    {
        /// <summary>
        /// Задает текст статусной строки
        /// </summary>
        /// <param name="bstrStatusLine">Текст статусной строки</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT SetStatusLine(BSTR bstrStatusLine);
        /// </prototype>
        /// </remarks>
        [PreserveSig]
        Moxel.HRESULT SetStatusLine( [MarshalAs(UnmanagedType.BStr)] string bstrStatusLine);

        /// <summary>
        /// Сброс статусной строки
        /// </summary>
        /// <remarks>
        /// <propotype>
        /// HRESULT ResetStatusLine();
        /// </propotype>
        /// </remarks>
        [PreserveSig]
        Moxel.HRESULT ResetStatusLine();
    }
}