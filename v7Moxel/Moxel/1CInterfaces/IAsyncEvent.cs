using System;
using System.Runtime.InteropServices;

namespace BMP1C.Net._1CInterfaces
{
    [ComVisible(true)]
    [Guid("ab634004-f13d-11d0-a459-004095e1daea")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAsyncEvent // : IUnknown
    {
        /// <summary>
        /// Установка глубины буфера событий
        /// </summary>
        /// <param name="depth">Buffer depth</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT SetEventBufferDepth(long lDepth);
        /// </prototype>
        /// </remarks>
        void SetEventBufferDepth(Int32 depth);

        /// <summary>
        /// Получение глубины буфера событий
        /// </summary>
        /// <param name="depth">Buffer depth</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetEventBufferDepth(long *plDepth);
        /// </prototype>
        /// </remarks>
        void GetEventBufferDepth(ref long depth);

        /// <summary>
        /// Посылка события
        /// </summary>
        /// <param name="source">Event source</param>
        /// <param name="message">Event message</param>
        /// <param name="data">Event data</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetEventBufferDepth(long *plDepth);
        /// </prototype>
        /// </remarks>
        void ExternalEvent(
            [MarshalAs(UnmanagedType.BStr)] String source,
            [MarshalAs(UnmanagedType.BStr)] String message,
            [MarshalAs(UnmanagedType.BStr)] String data
        );

        /// <summary>
        /// Очистка буфера событий
        /// </summary>
        /// <remarks>
        /// <prototype>
        /// </prototype>
        /// </remarks>
        void CleanBuffer();
    }
}