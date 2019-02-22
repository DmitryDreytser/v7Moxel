using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Ole
{
    [Flags]
    public enum OLEMISC
    {
        OLEMISC_RECOMPOSEONRESIZE = 0x1,
        OLEMISC_ONLYICONIC = 0x2,
        OLEMISC_INSERTNOTREPLACE = 0x4,
        OLEMISC_STATIC = 0x8,
        OLEMISC_CANTLINKINSIDE = 0x10,
        OLEMISC_CANLINKBYOLE1 = 0x20,
        OLEMISC_ISLINKOBJECT = 0x40,
        OLEMISC_INSIDEOUT = 0x80,
        OLEMISC_ACTIVATEWHENVISIBLE = 0x100,
        OLEMISC_RENDERINGISDEVICEINDEPENDENT = 0x200,
        OLEMISC_INVISIBLEATRUNTIME = 0x400,
        OLEMISC_ALWAYSRUN = 0x800,
        OLEMISC_ACTSLIKEBUTTON = 0x1000,
        OLEMISC_ACTSLIKELABEL = 0x2000,
        OLEMISC_NOUIACTIVATE = 0x4000,
        OLEMISC_ALIGNABLE = 0x8000,
        OLEMISC_SIMPLEFRAME = 0x10000,
        OLEMISC_SETCLIENTSITEFIRST = 0x20000,
        OLEMISC_IMEMODE = 0x40000,
        OLEMISC_IGNOREACTIVATEWHENVISIBLE = 0x80000,
        OLEMISC_WANTSTOMENUMERGE = 0x100000,
        OLEMISC_SUPPORTSMULTILEVELUNDO = 0x200000
    };

    public enum DVASPECT
    {
        DVASPECT_CONTENT = 1,
        DVASPECT_THUMBNAIL = 2,
        DVASPECT_ICON = 4,
        DVASPECT_DOCPRINT = 8
    }

    [ComImport, Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleObject
    {
        
        void SetClientSite([In, MarshalAs(UnmanagedType.Interface)] IOleClientSite pClientSite);
        void GetClientSite([MarshalAs(UnmanagedType.Interface), Out] out IOleClientSite ppClientSite);
        void SetHostNames([In, MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [In, MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);
        void Close([In, MarshalAs(UnmanagedType.U4)] uint dwSaveOption);
        void SetMoniker([In, MarshalAs(UnmanagedType.U4)] uint dwWhichMoniker, [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmk);
        void GetMoniker([In, MarshalAs(UnmanagedType.U4)] uint dwAssign, [In, MarshalAs(UnmanagedType.U4)] uint dwWhichMoniker, [MarshalAs(UnmanagedType.Interface), Out] out IMoniker ppmk);
        void InitFromData([In, MarshalAs(UnmanagedType.Interface)] IDataObject pDataObject, [In, MarshalAs(UnmanagedType.Bool)] bool fCreation, [In, MarshalAs(UnmanagedType.U4)] uint dwReserved);
        void GetClipboardData([In, MarshalAs(UnmanagedType.U4)] uint dwReserved, [MarshalAs(UnmanagedType.Interface), Out] out IDataObject ppDataObject);
        void DoVerb(OleVerbs iVerb, [In] IntPtr lpmsg, [In, MarshalAs(UnmanagedType.Interface)] IOleClientSite pActiveSite, int lindex, IntPtr hwndParent, [In] ref RECT lprcPosRect);
        void EnumVerbs([MarshalAs(UnmanagedType.Interface), Out] out IEnumOLEVERB ppEnumOleVerb);
        void Update();
        void IsUpToDate();
        void GetUserClassID([MarshalAs(UnmanagedType.Struct), Out] out Guid pClsid);
        void GetUserType([In, MarshalAs(UnmanagedType.U4)] uint dwFormOfType, [MarshalAs(UnmanagedType.LPWStr), Out] out string pszUserType);
        void SetExtent(uint dwDrawAspect, [MarshalAs(UnmanagedType.LPStruct)] tagSIZEL psizel);
        void GetExtent(uint dwDrawAspect, [MarshalAs(UnmanagedType.LPStruct), Out] tagSIZEL psizel);
        void Advise([In, MarshalAs(UnmanagedType.Interface)] IAdviseSink pAdvSink, [MarshalAs(UnmanagedType.U4), Out] out uint pdwConnection);
        void Unadvise([In, MarshalAs(UnmanagedType.U4)] uint dwConnection);
        void EnumAdvise([MarshalAs(UnmanagedType.Interface), Out] out IEnumStatData ppenumAdvise);
        void GetMiscStatus([In, MarshalAs(UnmanagedType.U4)] DVASPECT dwAspect, [MarshalAs(UnmanagedType.U4), Out] out OLEMISC pdwStatus);
        void SetColorScheme([In, MarshalAs(UnmanagedType.Struct)] object pLogpal);
    }

    [ComImport, Guid("00000104-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]

    public interface IEnumOLEVERB

    {
        [PreserveSig]
        HRESULT Next([MarshalAs(UnmanagedType.U4)] int celt, [Out] tagOLEVERB rgelt, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pceltFetched);
        [PreserveSig]
        HRESULT Skip([In, MarshalAs(UnmanagedType.U4)] int celt);
        void Reset();
        void Clone(out IEnumOLEVERB ppenum);
    }

    [StructLayout(LayoutKind.Sequential)]

    public sealed class tagOLEVERB
    {
        public int lVerb;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string lpszVerbName;
        [MarshalAs(UnmanagedType.U4)]
        public int fuFlags;
        [MarshalAs(UnmanagedType.U4)]
        public int grfAttribs;
        public tagOLEVERB() { }
    }
}