// ============================================================================
// PdfExport.cs — единая точка выхода PDF из перехвата SaveAs и из COM-объекта
// Converter. Пользователь выбирает, какой движок рисует PDF.
//
//   PdfEngine.Auto        — цепочка: векторный EMF -> растровый EMF -> Pechkin(HTML).
//                           Каждый шаг даёт fallback на следующий (как было в v2
//                           SaveWrapper, но с раздельными векторным/растровым EMF).
//   PdfEngine.PrintVector — PrintWrapper (перехват родной печати 1С -> EMF) +
//                           EmfVectorPdfWriter: выделяемый текст, встроенные TTF,
//                           точная печатная форма (колонтитулы &P/&Q, разрывы).
//   PdfEngine.PrintRaster — тот же захват печати, но EmfPdfWriter: страницы
//                           растеризуются (Dpi/Jpeg настраиваются через Settings).
//   PdfEngine.ModelHtml   — модель Moxel -> HtmlWriter -> Pechkin (wkhtmltopdf).
//                           Не требует ни view, ни принтера; качество — как у HTML.
//
// HeadlessPrintWrapper НЕ подключается (не протестирован).
//
// ЧТЕНИЕ ИЗ ПАМЯТИ: для ModelHtml используется быстрый прямой канал
// V7Table.LoadMoxelFromMemory() (MoxelMemLoader, без сериализации в mxl);
// при его неудаче — проверенный сериализационный ReadFromCSheetDoc.
// ============================================================================

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using static Moxel.MemoryReader;

namespace Moxel
{
    public static class PdfExport
    {
        /// <summary>Настройки растеризации для PrintRaster (Dpi, JPEG и т.п.). null — значения по умолчанию.</summary>
        public static EmfPdfSettings RasterSettings = null;

        /// <summary>Настройки вектора для PrintVector (шрифты, Compress). null — значения по умолчанию.</summary>
        public static EmfPdfSettings VectorSettings = null;

        /// <summary>Последняя ошибка (для диагностики из 1С).</summary>
        public static string LastError { get; private set; } = string.Empty;

        /// <summary>
        /// Сохранить таблицу (по указателю CSheetDoc) в PDF выбранным движком.
        /// table — опционально: управляемая обёртка (нужна для ModelHtml-канала
        /// и для печати никогда не показанной таблицы через Напечатать(0)).
        /// </summary>
        public static bool Save(IntPtr pSheetDoc, string fileName, PdfEngine engine, V7Table table = null)
        {
            LastError = string.Empty;

            switch (engine)
            {
                case PdfEngine.PrintVector:
                    return CapturePdf(pSheetDoc, fileName, false, table);
                case PdfEngine.PrintRaster:
                    return CapturePdf(pSheetDoc, fileName, true, table);
                case PdfEngine.ModelHtml:
                    return ModelHtmlPdf(pSheetDoc, fileName, table);
                default: // Auto: вектор -> растр -> HTML
                    if (CapturePdf(pSheetDoc, fileName, false, table)) return true;
                    string vecErr = LastError;
                    Debug.WriteLine($"PdfExport.Auto: вектор не удался ({vecErr}) — пробую растр");
                    if (CapturePdf(pSheetDoc, fileName, true, table)) return true;
                    string rasErr = LastError;
                    Debug.WriteLine($"PdfExport.Auto: растр не удался ({rasErr}) — fallback Pechkin");
                    if (ModelHtmlPdf(pSheetDoc, fileName, table)) return true;
                    LastError = $"Все каналы PDF исчерпаны. Вектор: {vecErr}; Растр: {rasErr}; HTML: {LastError}";
                    return false;
            }
        }

        /// <summary>Канал «перехват печати -> EMF -> PDF». raster=true — растеризация (EmfPdfWriter).</summary>
        static bool CapturePdf(IntPtr pSheetDoc, string fileName, bool raster, V7Table table)
        {
            string mode = raster ? "PrintRaster" : "PrintVector";
            try
            {
                if (!PrintWrapper.IsAvailable)
                {
                    LastError = $"{mode}: перехват печати недоступен ({PrintWrapper.LastError})";
                    return false;
                }

                var settings = raster ? RasterSettings : VectorSettings;

                if (PrintWrapper.CaptureToPdf(pSheetDoc, fileName, settings, table, raster))
                {
                    if (File.Exists(fileName) && new FileInfo(fileName).Length > 0)
                        return true;
                    LastError = $"{mode}: PDF не создан";
                    return false;
                }

                LastError = $"{mode}: {PrintWrapper.LastError}";
                return false;
            }
            catch (Exception ex)
            {
                LastError = $"{mode}: {ex.Message}";
                Debug.WriteLine($"PdfExport.{mode}: {ex}");
                return false;
            }
        }

        /// <summary>Канал «модель -> HTML -> Pechkin».</summary>
        static bool ModelHtmlPdf(IntPtr pSheetDoc, string fileName, V7Table table)
        {
            Converter.mxl?.Dispose();
            Converter.mxl = null;

            try
            {
                // Быстрый прямой канал чтения из памяти (без mxl-файла)
                if (table != null)
                {
                    try
                    {
                        Converter.mxl = table.LoadMoxelFromMemory();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"PdfExport.ModelHtml: MemLoader не дал модель ({ex.Message}) — сериализация");
                    }
                }

                // Fallback: сериализация через CArchive (проверенный путь)
                if (Converter.mxl == null)
                    Converter.mxl = ReadFromCSheetDoc(new CSheetDoc(pSheetDoc));

                if (Converter.mxl == null)
                {
                    LastError = "ModelHtml: не удалось прочитать таблицу";
                    return false;
                }

                ExcelWriter.PageSettings = new CSheetDoc(pSheetDoc).PageSettings;
                Converter.mxl.SaveAs(fileName, SaveFormat.PDF);
                return File.Exists(fileName);
            }
            catch (Exception ex)
            {
                Exception inner = ex;
                while (inner.InnerException != null) inner = inner.InnerException;
                LastError = $"ModelHtml: {inner.Message}";
                Debug.WriteLine($"PdfExport.ModelHtml: {ex}");
                return false;
            }
            finally
            {
                Converter.mxl?.Dispose();
                Converter.mxl = null;
            }
        }
    }
}
