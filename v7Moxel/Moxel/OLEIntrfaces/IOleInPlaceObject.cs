using System;
using System.Runtime.InteropServices;
using System.Linq;

namespace Ole
{
    [ComConversionLoss]
    [ComImport, Guid("00000113-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleInPlaceObject : IOleWindow
    {
        void InPlaceDeactivate();
        //
        void ReactivateAndUndo();
        //
        // Parameters:
        //   lprcPosRect:
        //
        //   lprcClipRect:
        void SetObjectRects(RECT[] lprcPosRect, RECT[] lprcClipRect);
        //
        void UIDeactivate();
    }
}
