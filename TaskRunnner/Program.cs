using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using v7Moxel;

namespace TaskRunnner
{
    static class Program
    {
        static bool ShowProgress = true;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// <param name="args">Параметры командной строки</param>
        /// первый параметр - имя файла mxl
        /// второй параметр - имя файла для сохранения
        /// третий параметр - указатель HWnd для возврата текущего прогресса конвертации
        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length == 0)
                return -1;

            string fileName = args[0];
            string resultFileName = args[1];
            
            if (args.Length == 3)
                ShowProgress = true;

            Debug.WriteLine(args[0]);
            Debug.WriteLine(args[1]);

            Moxel.Moxel mxl = new Moxel.Moxel(fileName);

            Debug.WriteLine("Загружена таблица");

            switch (Path.GetExtension(resultFileName).ToLower())
            {
                case ".xlsx":
                    Debug.WriteLine("Сохраняем в Excel");
                    if (ShowProgress)
                        Moxel.ExcelWriter.onProgress += ExcelWriter_onProgress;
                    mxl.SaveAs(resultFileName, Moxel.SaveFormat.Excel);
                    return 1;
                case ".html":
                    Debug.WriteLine("Сохраняем в Html");
                    if (ShowProgress)
                        Moxel.HtmlWriter.onProgress += ExcelWriter_onProgress;
                    mxl.SaveAs(resultFileName, Moxel.SaveFormat.Html);
                    return 1;
                case ".pdf":
                    Debug.WriteLine("Сохраняем в PDF");
                    if (ShowProgress)
                        Moxel.PDFWriter.onProgress += ExcelWriter_onProgress;
                    mxl.SaveAs(resultFileName, Moxel.SaveFormat.PDF);
                    return 1;
                default:
                    Debug.WriteLine("Формат не определен");
                    return -2;
            }

            //Debug.WriteLine("Сохраняем в Excel");
            //return 0;
        }

        private static void ExcelWriter_onProgress(int progress)
        {
            Console.WriteLine(progress);
        }
    }
}
