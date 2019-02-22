using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using IAdviseSink = System.Runtime.InteropServices.ComTypes.IAdviseSink;

namespace Ole
{
    [ComImport(), Guid("00000127-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IViewObject2 /* : IViewObject */
    {
        void Draw(
            [In, MarshalAs(UnmanagedType.U4)]
                int dwDrawAspect,

                 int lindex,

            IntPtr pvAspect,
            [In]
                ref DVTARGETDEVICE ptd,

            IntPtr hdcTargetDev,

            IntPtr hdcDraw,
            [In]
                ref COMRECT lprcBounds,
            [In]
                ref COMRECT lprcWBounds,

            IntPtr pfnContinue,
            [In]
                int dwContinue);


        [PreserveSig]
        int GetColorSet(
            [In, MarshalAs(UnmanagedType.U4)]
                int dwDrawAspect,

            int lindex,

            IntPtr pvAspect,
            [In]
                DVTARGETDEVICE ptd,

            IntPtr hicTargetDev,
            [Out]
                tagLOGPALETTE ppColorSet);


        [PreserveSig]
        int Freeze(
            [In, MarshalAs(UnmanagedType.U4)]
                int dwDrawAspect,

            int lindex,

            IntPtr pvAspect,
            [Out]
                IntPtr pdwFreeze);


        [PreserveSig]
        int Unfreeze(
            [In, MarshalAs(UnmanagedType.U4)]
                int dwFreeze);


        void SetAdvise(
            [In, MarshalAs(UnmanagedType.U4)]
                int aspects,
            [In, MarshalAs(UnmanagedType.U4)]
                int advf,
            [In, MarshalAs(UnmanagedType.Interface)]
                IAdviseSink pAdvSink);


        void GetAdvise(
            // These can be NULL if caller doesn't want them
            [In, Out, MarshalAs(UnmanagedType.LPArray)]
                int[] paspects,
            // These can be NULL if caller doesn't want them
            [In, Out, MarshalAs(UnmanagedType.LPArray)]
                int[] advf,
            // These can be NULL if caller doesn't want them
            [In, Out, MarshalAs(UnmanagedType.LPArray)]
                IAdviseSink[] pAdvSink);


        void GetExtent(
            [In, MarshalAs(UnmanagedType.U4)]
                int dwDrawAspect,

            int lindex,
            [In]
                DVTARGETDEVICE ptd,
            [Out]
                tagSIZEL lpsizel);
    }
}
