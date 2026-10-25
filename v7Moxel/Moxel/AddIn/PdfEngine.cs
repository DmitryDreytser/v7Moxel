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

using System.Runtime.InteropServices;

namespace Moxel
{
    /// <summary>Движок генерации PDF (COM-visible, индексы стабильны).</summary>
    [ComVisible(true)]
    public enum PdfEngine
    {
        /// <summary>Авто: печать+вектор -> печать+растр -> модель+HTML(Pechkin).</summary>
        Auto = 0,
        /// <summary>Перехват печати 1С -> EMF -> векторный PDF (текст выделяется).</summary>
        PrintVector = 1,
        /// <summary>Перехват печати 1С -> EMF -> растровый PDF (JPEG/Flate, Dpi).</summary>
        PrintRaster = 2,
        /// <summary>Модель MOXEL -> HTML -> Pechkin (wkhtmltopdf).</summary>
        ModelHtml = 3,
    }
}
