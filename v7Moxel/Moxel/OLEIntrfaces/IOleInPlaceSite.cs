using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [ComConversionLoss]
    [ComImport(), Guid("00000119-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleInPlaceSite
    {
        int CanInPlaceActivate();

        int DeactivateAndUndo();
        int DiscardUndoState();

        [PreserveSig]
        int GetWindowContext(
            [Out] out IntPtr ppFrame,
            [Out] out IntPtr ppDoc,
            [Out] out RECT lprcPosRect,
            [Out] out RECT lprcClipRect,
            [In, Out] IntPtr lpFrameInfo);

        int OnInPlaceActivate();

        int OnInPlaceDeactivate();

        int OnPosRectChange([In] RECT lprcPosRect);
        int OnUIActivate();
        int OnUIDeactivate([In, MarshalAs(UnmanagedType.Bool)] bool fUndoable);

        int Scroll([In] tagSIZEL scrollExtant);
    }
}
