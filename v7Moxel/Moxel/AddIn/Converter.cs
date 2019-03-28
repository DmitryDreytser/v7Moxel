using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static Moxel.MemoryReader;

namespace Moxel
{
    [ComVisible(false)]
    //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //[Guid("1EAE378F-C315-4B49-980C-A9A40792E78C")]
    internal interface IConverter
    {
        [Alias("Присоединить")]
        void Attach(object Table);

        [Alias("ПерехватВключен")]
        int IsWrapped { get;}

        [Alias("ЗагрузитьИзПамяти")]
        void ReadFromMemory(object Table);

        [Alias("Записать")]
        string Save(string filename, SaveFormat format);

        [Alias("ПерехватитьЗапись")]
        int WrapSaveAs(int doWrap = 1);

        [Alias("ОписаниеОшибки")]
        string GetErrorDescription();
    }


    [ComVisible(true)]
    [Guid("2DF0622D-BC0A-4C30-8B7D-ACB66FB837B6")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Description("Конвертер MOXEL")]
    [ProgId("AddIn.Moxel.Converter")]
    public class Converter : AddIn, IConverter
    {

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate void dAfxThrowOleDispatchException(int a1, [MarshalAs(UnmanagedType.LPStr)]string ErrorMessage, int Flag);
        public static dAfxThrowOleDispatchException AfxThrowOleDispatchException = MFCNative.GetDelegate<dAfxThrowOleDispatchException>(1268);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate IntPtr dGetExecutedModule();
        dGetExecutedModule GetExecutedModule = WinApi.GetDelegate<dGetExecutedModule>("blang.dll", "?GetExecutedModule@CBLModule@@SAPAV1@XZ");

        [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate void dCBLModule__Reset(IntPtr _this);
        dCBLModule__Reset OnRuntimeError = WinApi.GetDelegate<dCBLModule__Reset>("blang.dll", "?OnRuntimeError@CBLModule@@UAEHXZ");

        Moxel mxl;
        static int ObjectCount = 0;

        public int IsWrapped { get { return SaveWrapper.isWraped ? 1 : 0; } }

        public int WrapSaveAs(int doWrap = 1)
        {
            return SaveWrapper.Wrap(doWrap == 1);
        }


        public void ReadFromMemory(object Table)
        {
            try
            {
                CTableOutputContext TableObject = CObject.FromComObject<CTableOutputContext>(Table);
                mxl = ReadFromCSheetDoc(TableObject.SheetDoc);
            }
            catch (Exception ex)
            {
                while (Marshal.ReleaseComObject(Table) > 0) { }
                Marshal.FinalReleaseComObject(Table);
                throw ex.InnerException;
            }
        }



        public void Attach(object Table)
        {
            try
            {
                string tempfile = Path.GetTempFileName();
                File.Delete(tempfile);
                tempfile += ".mxl";
                object[] param = { tempfile, "mxl" };
                var tt = Table.GetType().InvokeMember("Write", BindingFlags.InvokeMethod, null, Table, param);

                if (File.Exists(tempfile))
                    mxl = new Moxel(tempfile);

                File.Delete(tempfile);

            }
            catch (Exception ex)
            {
                while (Marshal.ReleaseComObject(Table) > 0) { }
                Marshal.FinalReleaseComObject(Table);
                throw ex.InnerException;
            }
        }

        public string Save(string filename, SaveFormat format)
        {
            if (mxl != null)
            {
                mxl.SaveAs(filename, format);
                return filename;
            }
            else
            {
                throw new Exception("Таблица не загружена.");
            }
        }
        
        public string GetErrorDescription()
        {
            return this.ErrorDescription;
        }

        #region AddIn events

        protected override void OnInit()
        {
            
        }

        protected override HRESULT OnRegister()
        {
            ObjectCount++;
            WrapSaveAs(1);
            return HRESULT.S_OK;
        }

        protected override void OnDone()
        {
            if (--ObjectCount == 0) // При выгрузке последнего объекта из памяти отключим перехват, если он есть.
                SaveWrapper.Wrap(false);

            mxl = null;
        }

        #endregion

    }
}
