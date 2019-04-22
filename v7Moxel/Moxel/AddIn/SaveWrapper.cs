using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Moxel.MemoryReader;

namespace Moxel
{
    public static class SaveWrapper
    {

        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate void dRaiseExtRuntimeError([MarshalAs(UnmanagedType.LPStr)]string ErrorMessage, int Flag);
        public static dRaiseExtRuntimeError RaiseExtRuntimeError = WinApi.GetDelegate<dRaiseExtRuntimeError>("blang.dll", "?RaiseExtRuntimeError@CBLModule@@SAXPBDH@Z");

        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate int dSaveToExcel(IntPtr SheetDoc, IntPtr SheetGDI, string FileName);

        static dSaveToExcel WrapperSaveToExcel = new dSaveToExcel(SaveToExcel);
        static dSaveToExcel WrapperSaveToHtml = new dSaveToExcel(SaveToHtml);
        //static GCHandle wrapper = GCHandle.Alloc(WrapperSaveToExcel);
        static IntPtr pWrapperSaveToExcel = Marshal.GetFunctionPointerForDelegate<dSaveToExcel>(WrapperSaveToExcel);
        static IntPtr pWrapperSaveToHtml = Marshal.GetFunctionPointerForDelegate<dSaveToExcel>(WrapperSaveToHtml);

        public static bool isWraped = false;
        static IntPtr ProcRVASaveToExcel = IntPtr.Zero;
        static IntPtr ProcRVASaveToHtml = IntPtr.Zero;

        //SavetoExcel = 0x5B420
        //SavetoHtml = 0x5B760
        public static string utilPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TaskRunnner.exe");

        public static int SaveExternal(string MoxelName, string FileName)
        {
            using (Process ExCon = new Process())
            {
                ExCon.StartInfo = new ProcessStartInfo { FileName = utilPath, Arguments = $"\"{MoxelName}\" \"{FileName}\"", CreateNoWindow = true, RedirectStandardOutput = true, UseShellExecute = false };
                ExCon.OutputDataReceived += ExCon_OutputDataReceived;
                ExCon.Start();
                ExCon.BeginOutputReadLine();

                while (!ExCon.HasExited)
                {
                    ExCon.WaitForExit(500);
                    Application.RaiseIdle(EventArgs.Empty);
                }

                if (ExCon.ExitCode == 1 && File.Exists(FileName))
                    return ExCon.ExitCode;
                else
                    return 0;
            }

        }

        private static void ExCon_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            Converter.StatusLine($"{e.Data:D2}%");
        }

        static int SaveToExcel(IntPtr pSheetDoc, IntPtr SheetGDI, string FileName)
        {

            Converter.mxl = null;

            try
            {
                CSheetDoc SheetDoc = new CSheetDoc(pSheetDoc);
                if (SheetDoc.Length > 1024 * 1024 * 2) //Если примерный размер таблицы > 2Мб тогда конвертируем в отдельном процессе, чтобы не словить OutOfMemory
                {
                    string tmpFileName = Path.GetTempFileName();
                    SheetDoc.SaveToFile(tmpFileName);

                    if(SaveExternal(tmpFileName, FileName) == 1)
                    {
                        File.Delete(tmpFileName);
                        return 1;
                    }
                    else
                        RaiseExtRuntimeError?.Invoke($"Ошибка сохранения таблицы в XLSX.: ошибка внешнего конвертера", 0);
                }
                else
                {
                    Converter.mxl = ReadFromCSheetDoc(SheetDoc);
                    ExcelWriter.PageSettings = SheetDoc.PageSettings;
                }
            }
            catch (Exception ex)
            {
                RaiseExtRuntimeError?.Invoke($"Ошибка сохранения таблицы в XLSX.:{ex.Message}", 0);
                return 0;
            }

            if (Converter.mxl != null)
            {
                try
                {
                    if (FileName.EndsWith(".xlsx"))
                        Converter.mxl.SaveAs(FileName, SaveFormat.Excel);
                    else
                    {
                        if (FileName.EndsWith(".xls"))
                        {
                            Converter.mxl.SaveAs(FileName + "x", SaveFormat.Excel);
                            File.Move(FileName + "x", FileName);
                        }
                        else
                            Converter.mxl.SaveAs(FileName + ".xlsx", SaveFormat.Excel);
                    }
                    Converter.mxl = null;
                    GC.Collect();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return 1;
                }
                catch (Exception ex)
                {
                    Converter.mxl = null;
                    GC.Collect();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    RaiseExtRuntimeError?.Invoke($"Ошибка сохранения таблицы в XLSX: {ex.Message}", 0);
                }
            }
            else
            {
                Converter.mxl = null;
                GC.Collect();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                RaiseExtRuntimeError?.Invoke("Ошибка сохранения таблицы в XLSX. Не удалось прочитать таблицу", 0);
            }
            return 0;
        }

        static int SaveToHtml(IntPtr pSheetDoc, IntPtr SheetGDI, string FileName)
        {
            Converter.mxl = null;

            try
            {
                Converter.mxl = ReadFromMemory(pSheetDoc);
            }
            catch (Exception ex)
            {
                RaiseExtRuntimeError?.Invoke($"Ошибка сохранения таблицы в HTML :{ex.Message}", 0);
                return 0;
            }

            if (Converter.mxl != null)
            {
                Converter.mxl.SaveAs(FileName, SaveFormat.Html);
                return 1;
            }
            else
                RaiseExtRuntimeError?.Invoke("Ошибка сохранения таблицы в HTML. Не удалось прочитать таблицу", 0);

            return 0;
        }

        static int hMoxel = 0;
        static IntPtr hProcess;

        static byte[] OriginalBytesXls = new byte[6];
        static byte[] OriginalBytesHtml = new byte[6];

        static byte[] oldRes = new byte[288];

        static IntPtr FileSaveFilterResource;

        public static int Wrap(bool DoWrap)
        {

            if (hMoxel == 0)
            {
                hMoxel = WinApi.GetModuleHandle("Moxel.dll").ToInt32();
                ProcRVASaveToExcel = new IntPtr(hMoxel + 0x5B420);
                Marshal.Copy(ProcRVASaveToExcel, OriginalBytesXls, 0, 6);

                ProcRVASaveToHtml = new IntPtr(hMoxel + 0x5B760);
                Marshal.Copy(ProcRVASaveToHtml, OriginalBytesHtml, 0, 6);
            }

            uint Protection = 0;
            try
            {
                if (DoWrap)
                {
                    if (isWraped)
                        return 1;
                    hProcess = Process.GetCurrentProcess().Handle;

                    if (WinApi.VirtualProtectEx(hProcess, ProcRVASaveToExcel, new IntPtr(6), 0x40, out Protection))
                    {
                        #region Подмена сохранения в Excel
                        Debug.WriteLine($"Патч по адресу {ProcRVASaveToExcel.ToInt32():X8}: PUSH {pWrapperSaveToExcel.ToInt32():X8}; RET;");
                        Marshal.WriteByte(ProcRVASaveToExcel, 0x68); //PUSH
                        Marshal.WriteIntPtr(new IntPtr(ProcRVASaveToExcel.ToInt32() + 1), pWrapperSaveToExcel);
                        Marshal.WriteByte(new IntPtr(ProcRVASaveToExcel.ToInt32() + 5), 0xC3); //RET
                        WinApi.FlushInstructionCache(hProcess, ProcRVASaveToExcel, new IntPtr(6));
                        WinApi.VirtualProtectEx(hProcess, ProcRVASaveToExcel, new IntPtr(6), Protection, out Protection);
                        #endregion

                        #region Подмена сохранения в html

                        if (WinApi.VirtualProtectEx(hProcess, ProcRVASaveToHtml, new IntPtr(6), 0x40, out Protection))
                        {
                            Debug.WriteLine($"Патч по адресу {ProcRVASaveToHtml.ToInt32():X8}: PUSH {pWrapperSaveToHtml.ToInt32():X8}; RET;");
                            Marshal.WriteByte(ProcRVASaveToHtml, 0x68); //PUSH
                            Marshal.WriteIntPtr(new IntPtr(ProcRVASaveToHtml.ToInt32() + 1), pWrapperSaveToHtml);
                            Marshal.WriteByte(new IntPtr(ProcRVASaveToHtml.ToInt32() + 5), 0xC3); //RET
                            WinApi.FlushInstructionCache(hProcess, ProcRVASaveToHtml, new IntPtr(6));
                            WinApi.VirtualProtectEx(hProcess, ProcRVASaveToHtml, new IntPtr(6), Protection, out Protection);
                        }
                        #endregion


                        //Заменим фильтр диалога сохранения. Заменим *.xls на *.xlsx
                        #region Патч списка форматов
                        IntPtr hRCRus = WinApi.GetModuleHandle("1crcrus.dll");
                        IntPtr Res = WinApi.FindResource(hRCRus, WinApi.MakeIntResource(0x770), 6);
                        int len = WinApi.SizeofResource(hRCRus, Res);
                        Array.Resize(ref oldRes, len);
                        Res = WinApi.LoadResource(hRCRus, Res);
                        FileSaveFilterResource = WinApi.LockResource(Res);
                        Debug.WriteLine($"Ресурс найден по адресу {FileSaveFilterResource.ToInt32():X8}");
                        Marshal.Copy(FileSaveFilterResource, oldRes, 0, len);

                        Char[] newRes = new Char[256 / 2];
                        Array.Clear(newRes, 0, newRes.Length);

                        string[] DialogFilter = { "Таблица Excel 2007 (*.xlsx)|*.xlsx", "HTML Документ (*.html)|*.html", "Текстовый файл (*.txt)|*.txt" };

                        int charindex = 0;
                        foreach (string Filter in DialogFilter)
                        {
                            newRes[charindex++] = (Char)Filter.Length;

                            foreach (Char chr in Filter)
                            {
                                newRes[charindex++] = chr;
                            }
                        }

                        byte[] buffer = new byte[charindex * 2];

                        Array.Resize(ref newRes, charindex);

                        Buffer.BlockCopy(newRes, 0, buffer, 0, newRes.Length * 2);

                        if (WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(buffer.Length), 0x40, out Protection))
                        {

                            Marshal.Copy(buffer, 0, FileSaveFilterResource, buffer.Length);

                            WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(buffer.Length), Protection, out Protection);
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
                    if (WinApi.VirtualProtectEx(hProcess, ProcRVASaveToExcel, new IntPtr(6), 0x40, out Protection)) ;
                    {
                        Marshal.Copy(OriginalBytesXls, 0, ProcRVASaveToExcel, 6);
                        Debug.WriteLine($"Снят патч по адресу {ProcRVASaveToExcel.ToInt32():X8}");
                        WinApi.FlushInstructionCache(hProcess, ProcRVASaveToExcel, new IntPtr(6));
                        WinApi.VirtualProtectEx(hProcess, ProcRVASaveToExcel, new IntPtr(6), Protection, out Protection);
                    }

                    if (WinApi.VirtualProtectEx(hProcess, ProcRVASaveToHtml, new IntPtr(6), 0x40, out Protection)) ;
                    {
                        Marshal.Copy(OriginalBytesHtml, 0, ProcRVASaveToHtml, 6);
                        Debug.WriteLine($"Снят патч по адресу {ProcRVASaveToHtml.ToInt32():X8}");
                        WinApi.FlushInstructionCache(hProcess, ProcRVASaveToHtml, new IntPtr(6));
                        WinApi.VirtualProtectEx(hProcess, ProcRVASaveToHtml, new IntPtr(6), Protection, out Protection);
                    }

                    //Вернем фильтр диалога сохранения таблиц на место

                    if (WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(oldRes.Length), 0x40, out Protection)) ;
                    {
                        Marshal.Copy(oldRes, 0, FileSaveFilterResource, oldRes.Length);
                        WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(oldRes.Length), Protection, out Protection);
                    }

                    isWraped = false;
                }


                if (isWraped == DoWrap)
                    return 1;
                else
                    return 0;
            }
            catch
            {
                return 0;
            }

        }


    }
}
