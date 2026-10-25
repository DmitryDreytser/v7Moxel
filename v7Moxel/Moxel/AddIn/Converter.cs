using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using static Moxel.MemoryReader;


namespace Moxel
{


    [ComVisible(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1EAE378F-C315-4B49-980C-A9A40792E78C")]
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

        [Alias("ЗаписатьPDF")]
        string SavePdf(string filename, PdfEngine engine);

        [Alias("ПерехватитьЗапись")]
        int WrapSaveAs(int doWrap = 1);

        [Alias("ОписаниеОшибки")]
        string GetErrorDescription();

        [Alias("СтекОшибки")]
        string GetErrorStackTrace();

        [Alias("КоличествоПотоковКонвертации")]
        int MaxDegreeOfParallelism { get; set; }

        [Alias("ФорматPDF")]
        PdfEngine PdfEngine { get; set; }
    }


    [ComVisible(true)]    
    [Guid("2DF0622D-BC0A-4C30-8B7D-ACB66FB837B6")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComDefaultInterface(typeof(IConverter))]
    [Description("Конвертер MOXEL")]
    [ProgId( "AddIn.Moxel.Converter")]
    public class Converter : AddIn, IConverter
    {
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

        static Converter()
        {
            SaveWrapper.Wrap(true);
        }

        public Converter()
        {
            
        }

        public int IsWrapped { get => SaveWrapper.isWraped ? 1 : 0; }
        public int MaxDegreeOfParallelism { get => ExcelWriter.MaxDegreeOfParallelism ; set
            {
                ExcelWriter.MaxDegreeOfParallelism = value;
            }
        }
        private static V7Table _attachedTable = null;

        Moxel.PdfEngine _pdfEngine = Moxel.PdfEngine.Auto;

        public int WrapSaveAs(int doWrap = 1) => SaveWrapper.Wrap(doWrap == 1);
        
        public void ReadFromMemory(object Table) => Attach(Table);

        /// <summary>Движок PDF по умолчанию (используется в Save(...,PDF) и перехвате диалога).
        /// Присваивание — потокобезопасно (plain field write).</summary>
        public Moxel.PdfEngine PdfEngine
        {
            get { return _pdfEngine; }
            set { _pdfEngine = value; }
        }

        public void Attach(object Table)
        {
            mxl?.Dispose();
            _attachedTable?.Dispose();
            _attachedTable = null;
            try
            {
                _attachedTable = new V7Table();
                if (!_attachedTable.SetTable(Table))
                    throw new Exception(string.IsNullOrEmpty(_attachedTable.LastError)
                        ? "Не удалось получить CSheetDoc из таблицы."
                        : _attachedTable.LastError);
                PageSettings = _attachedTable.SheetDoc.PageSettings;
            }
            catch (Exception ex)
            {
                _attachedTable?.Dispose();
                _attachedTable = null;
                try
                {
                    while (Marshal.ReleaseComObject(Table) > 0) { }
                    Marshal.FinalReleaseComObject(Table);
                }
                catch { /* объект уже не RCW — игнорируем */ }
                // Бросаем исходное исключение с сохранением стека (раньше "throw ex.InnerException"
                // терял стек и давал NullReferenceException при отсутствии inner-ошибки).
                throw;
            }
        }

        public void Load(string FileNAme)
        {

            if (!File.Exists(FileNAme))
                throw new Exception($"Файл {FileNAme} не найден.");

            mxl?.Dispose();
            _attachedTable?.Dispose();
            _attachedTable = null;
            mxl = new Moxel(FileNAme);
        }

        public string Save(string filename, SaveFormat format)
        {
            try
            {
                if (_attachedTable != null)
                {
                    if (format == SaveFormat.PDF)
                        SaveWrapper.SaveToPDF(_attachedTable.Context.SheetDoc.Pointer, filename, _attachedTable);
                    else
                        SaveWrapper.WrapperSaveAs(_attachedTable.Context.SheetDoc.Pointer, filename, (int)format);

                    return filename;
                }

                else if (mxl == null)
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
            ObjectCount++;
        }

        protected override HRESULT OnRegister()
        {
            //ObjectCount++;
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
