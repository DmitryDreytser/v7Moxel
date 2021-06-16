using AddIn;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using v7Moxel.Moxel.ExcelWriter;
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

        [Alias("Загрузить")]
        void Load(string FileNAme);

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

        [Alias("СтекОшибки")]
        string GetErrorStackTrace();

        [Alias("КоличествоПотоковКонвертации")]
        int MaxDegreeOfParallelism { get; set; }
    }


    [ComVisible(true)]
    [Guid("2DF0622D-BC0A-4C30-8B7D-ACB66FB837B6")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(IConverter))]
    [Description("Конвертер MOXEL")]
    [ProgId("AddIn.Moxel.Converter")]
    public class Converter : AddIn, IConverter
    {

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate void dAfxThrowOleDispatchException(int a1, [MarshalAs(UnmanagedType.LPStr)]string ErrorMessage, int Flag);
        public static dAfxThrowOleDispatchException AfxThrowOleDispatchException = MFCNative.GetDelegate<dAfxThrowOleDispatchException>(1268);

        [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate void dCBLModule__Reset(IntPtr _this);
        dCBLModule__Reset OnRuntimeError = WinApi.GetDelegate<dCBLModule__Reset>("blang.dll", "?OnRuntimeError@CBLModule@@UAEHXZ");


        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Auto, ThrowOnUnmappableChar = true)]
        private delegate void dRaiseExtRuntimeError(IntPtr ErrorMessage, MessageMarker Flag);
        private static dRaiseExtRuntimeError RaiseExtRuntimeErrorNative = WinApi.GetDelegate<dRaiseExtRuntimeError>("blang.dll", "?RaiseExtRuntimeError@CBLModule@@SAXPBDH@Z");

        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        private delegate IntPtr dGetExecutedModule();
        
        private static dGetExecutedModule GetBkendUi = WinApi.GetDelegate<dGetExecutedModule>("bkend.dll", "?GetBkEndUI@@YAPAVCBkEndUI@@XZ");

        [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Auto, ThrowOnUnmappableChar = true)]
        private delegate void dDoMessageLine(IntPtr _this, IntPtr ErrorMessage, MessageMarker Flag);
        
        public enum MessageMarker
        {
            None = 0,
            BlueTriangle,
            Exclamation,
            Exclamation2,
            Exclamation3,
            Information,
            BlackErr,
            RedErr,
            MetaData,
            UnderlinedErr
        };

        public static void RaiseExtRuntimeError(string ErrorMessage)
        {
                var bkendUi = GetBkendUi();

                var vfTable = Marshal.ReadIntPtr(bkendUi);

                var DoMessageLine = Marshal.GetDelegateForFunctionPointer<dDoMessageLine>(Marshal.ReadIntPtr(vfTable + 0xC));

                DoMessageLine.Invoke(bkendUi, Marshal.StringToCoTaskMemAnsi(ErrorMessage), MessageMarker.RedErr);
            
        }

        public static Moxel mxl;
        public static PageSettings PageSettings = null;
        public static CTableOutputContext TableObject = null;
        static int ObjectCount = 0;

        public int IsWrapped { get => SaveWrapper.isWraped ? 1 : 0; }
        public int MaxDegreeOfParallelism { get => ExcelWriter.MaxDegreeOfParallelism ; set
            {
                ExcelWriter.MaxDegreeOfParallelism = value;
            }
        }
        public int WrapSaveAs(int doWrap = 1) => SaveWrapper.Wrap(doWrap == 1);
        public void ReadFromMemory(object Table)
        {
            try
            {
                TableObject = CObject.FromComObject<CTableOutputContext>(Table);
                PageSettings = TableObject.SheetDoc.PageSettings;

                while (Marshal.ReleaseComObject(Table) > 0) { }
                Marshal.FinalReleaseComObject(Table);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                mxl?.Dispose();
                mxl = null;
            }
        }



        public void Attach(object Table)
        {
            mxl?.Dispose();
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

                while (Marshal.ReleaseComObject(Table) > 0) { }
                Marshal.FinalReleaseComObject(Table);

            }
            catch (Exception ex)
            {
                while (Marshal.ReleaseComObject(Table) > 0) { }
                Marshal.FinalReleaseComObject(Table);
                throw ex.InnerException;
            }
        }

        public void Load(string FileNAme)
        {

            if (!File.Exists(FileNAme))
                throw new Exception($"Файл {FileNAme} не найден.");

            mxl?.Dispose();
            mxl = new Moxel(FileNAme);
        }

        public string Save(string filename, SaveFormat format)
        {
            try
            {
                if (mxl == null)
                {
                    if (TableObject != null)
                        if (TableObject.SheetDoc.Length < 1024 * 1024 * 2 || !SaveWrapper.CanSaveExternal)
                            mxl = ReadFromCSheetDoc(TableObject.SheetDoc);
                        else
                        {
                            string tmpFileName = Path.GetTempFileName();
                            TableObject.SheetDoc.SaveToFile(tmpFileName);
                            if (SaveWrapper.SaveExternal(tmpFileName, filename).Result == 1)
                            {
                                File.Delete(tmpFileName);
                                return filename;
                            }
                            else
                                throw new Exception("Ошибка записи.");
                        }
                    else
                        throw new Exception("Таблица не загружена.");

                }

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
            finally
            {
                mxl?.Dispose();
                mxl = null;
                GC.Collect();
                GC.Collect();
            }
        }


        #region AddIn events

        public string GetErrorDescription()
        {
            return this.ErrorDescription;
        }

        public string GetErrorStackTrace()
        {
            return this.ErrorStackTrace;
        }

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
