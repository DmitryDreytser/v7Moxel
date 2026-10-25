// SaveWrapper.cs (v2)
// Перехват CSheetDoc::SaveAs (moxel.dll) — патч 6 байт (PUSH stub; RET) + подмена списка форматов
// диалога сохранения (RT_STRING ресурс 1crcrus.dll).
//
// Что изменено относительно v1 (см. комменты «ФИКС»/«НОВОЕ»):
//   1. ФИКС: SaveAs обёрнут в try/catch — .NET-исключение, ушедшее в native черезmarshaled-делегат,
//      роняет процесс 1С без шансов на диагностику.
//   2. ФИКС: охрана GetModuleHandle/GetProcAddress == 0 (раньше Marshal.Copy из нулевого
//      адреса = AccessViolation при вызове Wrap до загрузки Moxel.dll).
//   3. НОВОЕ: SaveToPDF сначала пробует перехват родной печати (PrintWrapper → EMF → PDF),
//      при неудаче откатывается на Pechkin (HTML-путь). Печать даёт точную копию печатной
//      формы: колонтитулы с номерами страниц, разрывы, масштаб — то, чего HTML-путь не умеет.
//   4. ФИКС: SaveToHtml/SaveToPDF переведены на ReadFromCSheetDoc (прямой канал памяти,
//      тот же, что у Excel) вместо ReadFromMemory — сериализация в mxl для HTML/PDF не нужна.
//   5. ФИКС: буфер подмены фильтра диалога дозаполняется нулями до размера ресурса —
//      старые «призрачные» форматы после трёх новых больше не показываются.
//   6. Диагностика: Debug.WriteLine формата, пришедшего в SaveAs (калибровка отображения
//      индекса фильтра диалога на CSheetSaveAsType на живой 1С).
//
// Сохранившееся поведение v1: патч/снятие патчей, Resource-подмена, внешний конвертер
// (CanSaveExternal=false — ветка мёртвая, но оставлена для включения одной правкой).

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static Moxel.MemoryReader;

namespace Moxel
{
    public static class SaveWrapper
    {
        [UnmanagedFunctionPointer(CallingConvention.ThisCall, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate int dSaveAs(IntPtr SheetDoc, string FileName, int format);

        public static dSaveAs WrapperSaveAs = new dSaveAs(SaveAs);
        static IntPtr pWrapperSaveAs = Marshal.GetFunctionPointerForDelegate<dSaveAs>(WrapperSaveAs);

        public static bool isWraped = false;

        static IntPtr ProcSaveAs = IntPtr.Zero; // ФИКС: это VA (не RVA) — переименован для ясности

        //SavetoExcel = 0x5B420
        //SavetoHtml = 0x5B760
        //SaveToText = 0x2500A

        public static string utilPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TaskRunnner.exe");

        // ФИКС: убрано "&& false" из выражения — флаг теперь честный, ветка внешнего конвертера
        // включается установкой файла-утилиты. Поведение v1 (всегда false) не изменилось.
        public static bool CanSaveExternal => File.Exists(utilPath);

        // РЕФАКТОРИНГ (п.2): держатель RCW таблицы, присоединённой через COM-объект
        // Converter.Attach. Нужен прямому загрузчику MoxelMemLoader: пока жив RCW —
        // жива нативная таблица, и чтение по CSheetDoc* безопасно. Обновляется при
        // каждом Attach; при Detach/ошибке обнуляется Converter'ом.
        internal static object _lastAttachedComHolder = null;

        public static System.Threading.Tasks.Task<int> SaveExternal(string MoxelName, string FileName)
        {
            try
            {
                using (Process ExCon = new Process())
                {
                    ExCon.StartInfo = new ProcessStartInfo
                    {
                        FileName = utilPath,
                        Arguments = $"\"{MoxelName}\" \"{FileName}\" {ExcelWriter.MaxDegreeOfParallelism}",
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        UseShellExecute = false
                    };
                    ExCon.OutputDataReceived += ExCon_OutputDataReceived;
                    ExCon.Start();
                    ExCon.BeginOutputReadLine();

                    while (!ExCon.HasExited)
                    {
                        ExCon.WaitForExit(10);
                        Application.RaiseIdle(EventArgs.Empty);
                        Application.DoEvents();
                    }

                    if (ExCon.ExitCode == 1 && File.Exists(FileName))
                        return System.Threading.Tasks.Task.FromResult(ExCon.ExitCode);
                    return System.Threading.Tasks.Task.FromResult(0);
                }
            }
            catch (Exception e)
            {
                Debug.Fail($"Error: {e}");
                return System.Threading.Tasks.Task.FromResult(-1);
            }
        }

        static string progress;

        private static async void ExCon_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            Application.RaiseIdle(EventArgs.Empty);
            if (progress != e.Data)
            {
                await System.Threading.Tasks.Task.Factory.StartNew(() => { Converter.StatusLine($"{e.Data:D2}%"); });
                progress = e.Data;
            }
        }

        // ФИКС: единая точка входа с top-level try/catch — исключения не уходят в native.
        static int SaveAs(IntPtr pSheetDoc, string FileName, int format)
        {
            try
            {

                if (File.Exists(FileName))
                    File.Delete(FileName);

                Debug.WriteLine($"SaveWrapper.SaveAs: \"{FileName}\", format={format}, doc={pSheetDoc.ToInt32():X8}");

                switch (format)
                {
                    case 0:
                        new CSheetDoc(pSheetDoc).SaveToFile(FileName);
                        return 0;
                    case 1:
                        return SaveToExcel(pSheetDoc, FileName);
                    case 2:
                        return SaveToHtml(pSheetDoc, FileName);
                    case 3:
                        return SaveToPDF(pSheetDoc, FileName);
                    default:
                        return 1;
                }
            }
            catch (Exception ex)
            {
                Debug.Fail($"SaveAs: необработанное исключение: {ex}");
                Converter.RaiseExtRuntimeError($"Ошибка сохранения таблицы: {ex.Message}");
                return 0;
            }
        }

        static int SaveToExcel(IntPtr pSheetDoc, string FileName)
        {
            Converter.mxl?.Dispose();

            try
            {
                CSheetDoc SheetDoc = new CSheetDoc(pSheetDoc);
                if (SheetDoc.Length > 1024 * 1024 * 2 && CanSaveExternal) //Если примерный размер таблицы > 2Мб тогда конвертируем в отдельном процессе, чтобы не словить OutOfMemory
                {
                    string tmpFileName = Path.GetTempFileName();
                    SheetDoc.SaveToFile(tmpFileName);

                    if (SaveExternal(tmpFileName, FileName).Result == 1)
                    {
                        File.Delete(tmpFileName);
                        return 1;
                    }
                    Converter.RaiseExtRuntimeError("Ошибка сохранения таблицы в XLSX.: ошибка внешнего конвертера");
                }
                else
                {
                    // РЕФАКТОРИНГ (п.2): единая точка загрузки модели — быстрый прямой
                    // канал MoxelMemLoader с fallback на сериализацию (был только он).
                    Converter.mxl = LoadModel(pSheetDoc, _lastAttachedComHolder);
                    ExcelWriter.PageSettings = SheetDoc.PageSettings;
                }
            }
            catch (Exception ex)
            {
                Debug.Fail($"Error: {ex}");
                return 0;
            }

            if (Converter.mxl != null)
            {
                try
                {
                    if (FileName.EndsWith(".xlsx"))
                    {
                        Converter.mxl.SaveAs(FileName, SaveFormat.Excel);
                    }
                    else if (FileName.EndsWith(".xls"))
                    {
                        // ВНИМАНИЕ (поведение v1): в файл с расширением .xls пишется XLSX-контент.
                        // Excel откроет его с предупреждением «формат не соответствует расширению».
                        // Если предупреждение мешает — сохранять в настоящий xls/xlsx.
                        Converter.mxl.SaveAs(FileName + "x", SaveFormat.Excel);
                        File.Move(FileName + "x", FileName);
                    }
                    else
                    {
                        Converter.mxl.SaveAs(FileName + ".xlsx", SaveFormat.Excel);
                    }
                    return 1;
                }
                catch (Exception ex)
                {
                    Converter.RaiseExtRuntimeError($"Ошибка сохранения таблицы в XLSX: {ex}");
                }
                finally
                {
                    Converter.mxl?.Dispose();
                    Converter.mxl = null;
                }
            }
            else
            {
                Converter.RaiseExtRuntimeError("Ошибка сохранения таблицы в XLSX. Не удалось прочитать таблицу");
            }
            return 0;
        }

        static int SaveToHtml(IntPtr pSheetDoc, string FileName)
        {
            Converter.mxl = null;

            try
            {
                // РЕФАКТОРИНГ (п.2): единая точка загрузки модели — быстрый прямой
                // канал MoxelMemLoader с fallback на сериализацию (был только он).
                Converter.mxl = LoadModel(pSheetDoc, _lastAttachedComHolder);
            }
            catch (Exception ex)
            {
                Converter.RaiseExtRuntimeError($"Ошибка сохранения таблицы в HTML :{ex.Message}");
                return 0;
            }

            if (Converter.mxl != null)
            {
                try
                {
                    Converter.mxl.SaveAs(FileName, SaveFormat.Html);
                    return 1;
                }
                catch (Exception ex)
                {
                    Converter.RaiseExtRuntimeError($"Ошибка сохранения таблицы в HTML: {ex.Message}");
                }
                finally
                {
                    Converter.mxl.Dispose();
                    Converter.mxl = null;
                }
            }
            else
                Converter.RaiseExtRuntimeError("Ошибка сохранения таблицы в HTML. Не удалось прочитать таблицу");

            Converter.RaiseExtRuntimeError($"Ошибка сохранения таблицы в PDF: {PdfExport.LastError}");
            return 0;
        }

        // РЕФАКТОРИНГ (п.1): весь вывод PDF идёт через единую точку PdfExport.Save
        // (цепочка движков и fallback-и реализованы там). Этот метод — тонкая обёртка:
        // конвертирует bool в коды 1/0, принятые у перехвата SaveAs, и поднимает ошибку
        // в UI 1С. Дублирующая логика PrintWrapper/ReadFromCSheetDoc/Pechkin отсюда удалена.
        public static int SaveToPDF(IntPtr pSheetDoc, string FileName, V7Table table = null, PdfEngine engine = PdfEngine.Auto)
        {
            if (PdfExport.Save(pSheetDoc, FileName, engine, table))
                return 1;

            Converter.RaiseExtRuntimeError($"Ошибка сохранения таблицы в PDF: {PdfExport.LastError}");
            return 0;
        }

        static int hMoxel = 0;
        static IntPtr hProcess;

        static byte[] OriginalBytesSaveAs = new byte[6];

        static byte[] oldRes = new byte[288];

        static IntPtr FileSaveFilterResource;

        /// <summary>Инициализация адресов. Идемпотентна и вызывается ЛЕНИВО при первом
        /// Wrap(true), а НЕ из статического конструктора класса: раньше static-ctor
        /// SaveWrapper тянул за собой делегат WrapperSaveAs → GetFunctionPointerForDelegate,
        /// что приводило к инициализации типа Converter, который в своём static-ctor вызывал
        /// Wrap(true). Патч CSheetDoc::SaveAs ставился при первой же ссылке на любой тип
        /// сборки — ещё до подключения аддина к сеансу 1С и без возможности снятия.</summary>
        private static bool Init()
        {
            if (hMoxel != 0)
                return ProcSaveAs != IntPtr.Zero;

            IntPtr hMod = WinApi.GetModuleHandle("Moxel.dll");
            if (hMod == IntPtr.Zero)
            {
                Debug.WriteLine("SaveWrapper.Wrap: Moxel.dll не загружен");
                return false;
            }
            hMoxel = hMod.ToInt32();
            ProcSaveAs = WinApi.GetProcAddress(hMod, "?SaveAs@CSheetDoc@@QAEHPBDW4CSheetSaveAsType@@@Z");
            if (ProcSaveAs == IntPtr.Zero)
            {
                Debug.WriteLine("SaveWrapper.Wrap: экспорт CSheetDoc::SaveAs не найден");
                hMoxel = 0;
                return false;
            }
            Marshal.Copy(ProcSaveAs, OriginalBytesSaveAs, 0, 6);
            return true;
        }

        public static int Wrap(bool DoWrap)
        {
            // ФИКС: охрана адресов — раньше Copy из нулевого указателя = AV процесса 1С
            if (DoWrap && !Init())
                return 0;
            if (hMoxel == 0 && DoWrap)
            {
                IntPtr hMod = WinApi.GetModuleHandle("Moxel.dll");
                if (hMod == IntPtr.Zero)
                {
                    Debug.WriteLine("SaveWrapper.Wrap: Moxel.dll не загружен");
                    return 0;
                }
                hMoxel = hMod.ToInt32();
                ProcSaveAs = WinApi.GetProcAddress(hMod, "?SaveAs@CSheetDoc@@QAEHPBDW4CSheetSaveAsType@@@Z");
                if (ProcSaveAs == IntPtr.Zero)
                {
                    Debug.WriteLine("SaveWrapper.Wrap: экспорт CSheetDoc::SaveAs не найден");
                    hMoxel = 0;
                    return 0;
                }
                Marshal.Copy(ProcSaveAs, OriginalBytesSaveAs, 0, 6);
            }

            uint Protection = 0;
            try
            {
                if (DoWrap)
                {
                    if (isWraped)
                        return 1;
                    hProcess = Process.GetCurrentProcess().Handle;

                    if (WinApi.VirtualProtectEx(hProcess, ProcSaveAs, new IntPtr(6), 0x40, out Protection))
                    {
                        #region Подмена сохранения в SaveAs
                        Debug.WriteLine($"Патч по адресу {ProcSaveAs.ToInt32():X8}: PUSH {pWrapperSaveAs.ToInt32():X8}; RET;");
                        Marshal.WriteByte(ProcSaveAs, 0x68); //PUSH
                        Marshal.WriteIntPtr(new IntPtr(ProcSaveAs.ToInt32() + 1), pWrapperSaveAs);
                        Marshal.WriteByte(new IntPtr(ProcSaveAs.ToInt32() + 5), 0xC3); //RET
                        WinApi.FlushInstructionCache(hProcess, ProcSaveAs, new IntPtr(6));
                        WinApi.VirtualProtectEx(hProcess, ProcSaveAs, new IntPtr(6), Protection, out Protection);
                        #endregion

                        //Заменим фильтр диалога сохранения. Заменим *.xls на *.xlsx
                        #region Патч списка форматов
                        IntPtr hRCRus = WinApi.GetModuleHandle("1crcrus.dll");
                        if (hRCRus != IntPtr.Zero)
                        {
                            IntPtr Res = WinApi.FindResource(hRCRus, WinApi.MakeIntResource(0x770), 6);
                            int len = WinApi.SizeofResource(hRCRus, Res);
                            Array.Resize(ref oldRes, len);
                            Res = WinApi.LoadResource(hRCRus, Res);
                            FileSaveFilterResource = WinApi.LockResource(Res);
                            Debug.WriteLine($"Ресурс найден по адресу {FileSaveFilterResource.ToInt32():X8}");
                            Marshal.Copy(FileSaveFilterResource, oldRes, 0, len);

                            Char[] newRes = new Char[256 / 2];
                            Array.Clear(newRes, 0, newRes.Length);

                            string[] DialogFilter = { "Таблица Excel 2007 (*.xlsx)|*.xlsx", "HTML Документ (*.html)|*.html", "PDF (*.pdf)|*.pdf" };

                            int charindex = 0;
                            foreach (string Filter in DialogFilter)
                            {
                                newRes[charindex++] = (Char)Filter.Length;

                                foreach (Char chr in Filter)
                                {
                                    newRes[charindex++] = chr;
                                }
                            }

                            // ФИКС: дозаполняем буфер нулями до размера ресурса — хвост старых записей
                            // больше не «протекает» в диалог как призрачные форматы
                            int resChars = len / 2;
                            if (charindex < resChars)
                                Array.Resize(ref newRes, resChars);

                            byte[] buffer = new byte[newRes.Length * 2];
                            Buffer.BlockCopy(newRes, 0, buffer, 0, newRes.Length * 2);

                            if (WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(buffer.Length), 0x40, out Protection))
                            {
                                Marshal.Copy(buffer, 0, FileSaveFilterResource, buffer.Length);
                                WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(buffer.Length), Protection, out Protection);
                            }
                        }
                        else
                        {
                            Debug.WriteLine("SaveWrapper.Wrap: 1crcrus.dll не загружен — фильтр диалога не подменён");
                        }
                        #endregion

                        isWraped = true;
                    }
                }
                else
                {
                    if (!isWraped)
                        return 1;

                    hProcess = Process.GetCurrentProcess().Handle;

                    if (WinApi.VirtualProtectEx(hProcess, ProcSaveAs, new IntPtr(6), 0x40, out Protection))
                    {
                        Marshal.Copy(OriginalBytesSaveAs, 0, ProcSaveAs, 6);
                        Debug.WriteLine($"Снят патч по адресу {ProcSaveAs.ToInt32():X8}");
                        WinApi.FlushInstructionCache(hProcess, ProcSaveAs, new IntPtr(6));
                        WinApi.VirtualProtectEx(hProcess, ProcSaveAs, new IntPtr(6), Protection, out Protection);
                    }

                    //Вернем фильтр диалога сохранения таблиц на место
                    if (FileSaveFilterResource != IntPtr.Zero &&
                        WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(oldRes.Length), 0x40, out Protection))
                    {
                        Marshal.Copy(oldRes, 0, FileSaveFilterResource, oldRes.Length);
                        WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(oldRes.Length), Protection, out Protection);
                    }

                    isWraped = false;
                }

                if (isWraped == DoWrap)
                    return 1;
                return 0;
            }
            catch
            {
                return 0;
            }
        }
    }
}
