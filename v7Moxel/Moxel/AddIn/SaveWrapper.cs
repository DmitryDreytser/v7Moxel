using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static Moxel.MemoryReader;

namespace Moxel
{
    public static class SaveWrapper
    {
        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate int dSaveToExcel(IntPtr SheetDoc, IntPtr SheetGDI, string FileName);

        static dSaveToExcel WrapperProc = new dSaveToExcel(SaveToExcel);
        static GCHandle wrapper = GCHandle.Alloc(WrapperProc);
        static IntPtr pWrapperProc = Marshal.GetFunctionPointerForDelegate<dSaveToExcel>(WrapperProc);

        static bool isWraped = false;
        static IntPtr ProcRVA = IntPtr.Zero;

        static int SaveToExcel(IntPtr pSheetDoc, IntPtr SheetGDI, string FileName)
        {
            CFile f = CFile.FromHFile(IntPtr.Zero);
            int result = 0;
            try
            {
                CSheetDoc SheetDoc = new CSheetDoc(pSheetDoc);
                
                CArchive Arch = new CArchive(f, SheetDoc);
                SheetDoc.Serialize(Arch);
                Arch.Flush();
                Arch = null;
                byte[] buffer = f.GetBufer();
                f = null;

                Moxel mxl = new Moxel(ref buffer);

                if (FileName.EndsWith(".xlsx"))
                    mxl.SaveAs(FileName, SaveFormat.Excel);
                else
                {
                    if (FileName.EndsWith(".xls"))
                    {
                        mxl.SaveAs(FileName + "x", SaveFormat.Excel);
                        File.Move(FileName + "x", FileName);
                    }
                    else
                        mxl.SaveAs(FileName + ".xlsx", SaveFormat.Excel);
                }
                result = 1;
            }
            catch(Exception ex)
            {
                f.unpatch(); // Обязательно снять перехват
                f = null;
                result = 0;
            }

            return result;
        }

        static int hMoxel = 0;
        static IntPtr hProcess;// = Process.GetCurrentProcess().Handle;

        static byte[] OriginalBytes = new byte[6];

        static byte[] oldRes = new byte[288];

        static IntPtr FileSaveFilterResource;

        public static int Wrap(bool DoWrap)
        {

            if (hMoxel == 0)
            {
                hMoxel = WinApi.GetModuleHandle("Moxel.dll").ToInt32();
                ProcRVA = new IntPtr(hMoxel + 0x5B420);
                Marshal.Copy(ProcRVA, OriginalBytes, 0, 6);
            }

            uint Protection = 0;
            try
            {
                if (DoWrap)
                {
                    if (isWraped)
                        return 1;
                    hProcess = Process.GetCurrentProcess().Handle;

                    IntPtr RetAddress = Marshal.GetFunctionPointerForDelegate<dSaveToExcel>(WrapperProc);

                    if (WinApi.VirtualProtectEx(hProcess, ProcRVA, new IntPtr(6), 0x40, out Protection))
                    {
                        Debug.WriteLine($"Патч по адресу {ProcRVA.ToInt32():X8}: PUSH {RetAddress.ToInt32():X8}; RET;");

                        Marshal.WriteByte(ProcRVA, 0x68); //PUSH
                        Marshal.WriteIntPtr(new IntPtr(ProcRVA.ToInt32() + 1), RetAddress);
                        Marshal.WriteByte(new IntPtr(ProcRVA.ToInt32() + 5), 0xC3); //RET

                        WinApi.FlushInstructionCache(hProcess, ProcRVA, new IntPtr(6));

                        WinApi.VirtualProtectEx(hProcess, ProcRVA, new IntPtr(6), Protection, out Protection);
                        

                        //Заменим фильтр диалога сохранения. Заменим *.xls на *.xlsx

                        IntPtr hRCRus = WinApi.GetModuleHandle("1crcrus.dll");

                        IntPtr Res = WinApi.FindResource(hRCRus, WinApi.MakeIntResource(0x770), 6);

                        int len = WinApi.SizeofResource(hRCRus, Res);
                        Array.Resize(ref oldRes, len);

                        Res = WinApi.LoadResource(hRCRus, Res);

                        FileSaveFilterResource = WinApi.LockResource(Res);

                        Debug.WriteLine($"Ресурс найден по адресу {FileSaveFilterResource.ToInt32():X8}");

                        Marshal.Copy(FileSaveFilterResource, oldRes, 0, len);


                        Char[] newRes = new Char[len / 2];
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

                        isWraped = true;
                    }

                }
                else
                {
                    if (!isWraped)
                        return 1;

                    hProcess = Process.GetCurrentProcess().Handle;

                    if (WinApi.VirtualProtectEx(hProcess, ProcRVA, new IntPtr(6), 0x40, out Protection));
                    {
                        Marshal.Copy(OriginalBytes, 0, ProcRVA, 6);

                        Debug.WriteLine($"Снят патч по адресу {ProcRVA.ToInt32():X8}");

                        WinApi.FlushInstructionCache(hProcess, ProcRVA, new IntPtr(6));

                        WinApi.VirtualProtectEx(hProcess, ProcRVA, new IntPtr(6), Protection, out Protection);


                        //Вернем фильтр диалога сохранения таблиц на место

                        if(WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(oldRes.Length), 0x40, out Protection));
                        {

                            Marshal.Copy(oldRes, 0, FileSaveFilterResource, oldRes.Length);

                            WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(oldRes.Length), Protection, out Protection);
                        }

                        isWraped = false;
                    }
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
