// ============================================================================
//  ExcelWriter.cs — экспорт таблицы 1С 7.7 (модель Moxel) в XLSX (EPPlus 5..7)
//
//  Версия 2.0 — рефакторинг под прямое чтение таблицы из памяти
//  (V7Table.SetTable → MemSheet → MoxelMemLoader → Moxel → ExcelWriter.Save).
//
//  Что изменилось по сравнению с исходной версией:
//
//  КОРРЕКТНОСТЬ РАБОТЫ С ТАБЛИЦЕЙ
//  1. Введён ЭФФЕКТИВНЫЙ формат ячейки: умолчание листа ← колонка ← строка ←
//     ячейка (наложение по битам маски MoxelCellFlags). Раньше использовался
//     только FormatCell самой ячейки, из-за чего пропадали границы/фон/шрифты,
//     заданные на колонке или строке. Для таблиц из памяти это критично:
//     MoxelMemLoader копирует форматы 1:1 (дельты), наложение — задача рендера.
//  2. Исправлен copy-paste баг: заливка проверяла BorderColor != White вместо
//     BgColor != White (белый фон теперь корректно не рисуется).
//  3. Числа: сначала берётся родное значение ячейки (бит Value — результат
//     выражения 1С, в памяти это m_strDetails), затем — прежняя эвристика по
//     точкам/запятым в тексте. Целые → "#,##0", дробные → "#,##0.00".
//  4. Авто-объединение «по выделению»: добавлены границы цикла (был риск
//     бесконечного цикла на неполных данных) и проверка пересечения с Unions.
//  5. Шрифт: FontList через TryGetValue (был прямой индексатор → падение),
//     размер |wFontSize|/4 (был -wFontSize/4 → отрицательный размер → AV GDI+).
//  6. Колонтитулы пишутся в Odd* (при differentOddEven=false в Excel видны
//     именно они; Even* в старой версии не отображались вообще).
//  7. Поворот текста нормализуется (1С 270° → Excel 90°, диапазон 0..180/255),
//     иначе EPPlus бросает ArgumentOutOfRangeException.
//  8. Ориентация страницы маппится по значению 1С (1→портрет, 2→ландшафт)
//     без привязки к нумерации enum eOrientation; guard'ы на Paper/Scale.
//  9. Убран deadlock: FinishEvent.WaitOne() навсегда ждал при nAllRowCount=0
//     (AutoResetEvent к тому же не Resetился); Parallel.For и так синхронен.
// 10. MaxDegreeOfParallelism = Processors/4*3 даёт 0 на 1-2 ядрах → исключение.
//     Теперь минимум 1 (важно для старых ВМ под 1С).
// 11. Убрана мутация модели при экспорте (row.Height больше не пишется назад).
// 12. Объекты: незагруженные картинки/OLE (pObject == null, типично для
//     memory-пути v1) пропускаются вместо экспорта заглушек 1x1; guard'ы от
//     деления на ноль по ImageArea; Rows.TryGetValue для текстовых блоков.
//
//  ПРОИЗВОДИТЕЛЬНОСТЬ
//  13. Двухфазный конвейер: фаза 1 — Parallel.For над ЧИСТЫМИ вычислениями
//      (эффективные форматы, числа, измерение высот GDI), фаза 2 — быстрый
//      однопоточный перенос в EPPlus. Убраны гонки EPPlus (его внутренние
//      структуры не потокобезопасны) и lock'и на каждую строку.
//  14. Кэш GDI-шрифтов: был new Font() на каждую ячейку (десятки тысяч
//      GDI-хэндлов) — стало несколько десятков на таблицу.
//  15. Пустые ячейки без стилей не создают ExcelRange вообще (раньше — на
//      каждую ячейку таблицы, включая полностью пустые).
//  16. Карта объединений вместо чтения worksheet.MergedCells на каждую ячейку.
//  17. Подсчёт '.'/',' без ToCharArray()+LINQ; per-thread Graphics для измерений.
//
//  СОВМЕСТИМОСТЬ
//  Публичный API сохранён: ExcelWriter.Save(moxel, filename), OnProgress,
//  PageSettings, MaxDegreeOfParallelism; расширения MoxelExtentions оставлены
//  (SetHorisonttalAlign/SetVerticalAlign/SetBorder/GetFont — используются и
//  другими писателями). Целевой EPPlus — 5..7 (LicenseContext).
// ============================================================================

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Moxel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using static Moxel.Moxel;

namespace Moxel
{
    /// <summary>
    /// Перенос шрифтов/выравнивания/рамок из модели Moxel в EPPlus.
    /// Старые перегрузки с System.Drawing.Font сохранены для совместимости
    /// (их вызывают HtmlWriter/PDFWriter и код объектов).
    /// </summary>
    internal static class MoxelExtentions
    {
        public static void SetFromFont(this ExcelFont excelFont, Font f)
        {
            if (excelFont != null)
            {
                using (f)
                {
                    excelFont.SetFromFont(f.Name, f.Size, f.Bold, f.Italic, f.Underline, f.Strikeout);
                }
            }
        }

        public static void SetFromFont(this ExcelParagraph excelParagraph, Font f)
        {
            if (excelParagraph != null)
            {
                using (f)
                {
                    excelParagraph.SetFromFont(f.Name, f.Size, f.Bold, f.Italic, f.Underline, f.Strikeout);
                }
            }
        }

        /// <summary>Установка шрифта без создания System.Drawing.Font (без GDI-хэндлов).</summary>
        public static void SetFromFont(this ExcelFont excelFont, string name, float size, bool bold, bool italic, bool underline)
        {
            if (excelFont == null) return;
            if (string.IsNullOrEmpty(name)) name = "Arial";
            if (size < 1f) size = 8f;
            excelFont.SetFromFont(name, size, bold, italic, underline, false);
        }

        /// <summary>Установка шрифта абзаца без создания System.Drawing.Font.</summary>
        public static void SetFromFont(this ExcelParagraph excelParagraph, string name, float size, bool bold, bool italic, bool underline)
        {
            if (excelParagraph == null) return;
            if (string.IsNullOrEmpty(name)) name = "Arial";
            if (size < 1f) size = 8f;
            excelParagraph.SetFromFont(name, size, bold, italic, underline, false);
        }

        /// <summary>
        /// Шрифт ячейки (System.Drawing) — используется для объектов-текстов.
        /// Шрифт наследуется: ячейка → умолчание. wFontSize в модели знаковый
        /// (-4 * размер_пт), поэтому модуль делится на 4.
        /// </summary>
        public static Font GetFont(this DataCell moxelCell, string defFontName, int defFontSize)
        {
            var flags = moxelCell.FormatCell.dwFlags;

            var name = defFontName;
            if (flags.HasFlag(MoxelCellFlags.FontName))
            {
                LOGFONT lf;
                if (moxelCell.Parent != null && moxelCell.Parent.FontList != null &&
                    moxelCell.Parent.FontList.TryGetValue(moxelCell.FormatCell.wFontNumber, out lf) &&
                    !string.IsNullOrEmpty(lf.lfFaceName))
                    name = lf.lfFaceName;
            }

            var size = defFontSize;
            if (flags.HasFlag(MoxelCellFlags.FontSize))
                size = Math.Abs(moxelCell.FormatCell.wFontSize) / 4;   // знаковый: -4*пт
            if (size < 1) size = defFontSize < 1 ? 8 : defFontSize;

            var fontStyle = FontStyle.Regular;
            if (flags.HasFlag(MoxelCellFlags.FontWeight) && moxelCell.FormatCell.bFontBold == clFontWeight.Bold)
                fontStyle |= FontStyle.Bold;
            if (flags.HasFlag(MoxelCellFlags.FontItalic) && moxelCell.FormatCell.bFontItalic)
                fontStyle |= FontStyle.Italic;
            if (flags.HasFlag(MoxelCellFlags.FontUnderline) && moxelCell.FormatCell.bFontUnderline)
                fontStyle |= FontStyle.Underline;

            return new Font(string.IsNullOrEmpty(name) ? defFontName : name, size, fontStyle);
        }

        public static void SetHorisonttalAlign(this DataCell moxelCell, MoxelRow row, ExcelStyle style)
        {
            var moxel = moxelCell.Parent;

            var alignHor =
              moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH)
                  ? moxelCell.FormatCell.bHorAlign
                  : (row != null && row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignH)
                      ? row.FormatCell.bHorAlign
                      : (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.AlignH)
                          ? moxel.DefFormat.bHorAlign
                          : TextHorzAlign.Left));

            switch (alignHor)
            {
                case TextHorzAlign.Left:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    break;
                case TextHorzAlign.Right:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    break;
                case TextHorzAlign.Center:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    break;
                case TextHorzAlign.Justify:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                    break;
            }
        }

        public static void SetVerticalAlign(this DataCell moxelCell, MoxelRow row, ExcelStyle style)
        {
            var moxel = moxelCell.Parent;

            var alignVert =
              moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV)
                  ? moxelCell.FormatCell.bVertAlign
                  : (row != null && row.FormatCell.dwFlags.HasFlag(MoxelCellFlags.AlignV)
                      ? row.FormatCell.bVertAlign
                      : (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.AlignV)
                          ? moxel.DefFormat.bVertAlign
                          : TextVertAlign.Bottom));

            switch (alignVert)
            {
                case TextVertAlign.Bottom:
                    style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                    break;
                case TextVertAlign.Middle:
                    style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    break;
                case TextVertAlign.Top:
                    style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    break;
            }
        }

        public static void SetBorder(this DataCell moxelCell, Border border)
        {
            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderBottom))
            {
                border.Bottom.Style = ExcelWriter.GetBorderStyle(moxelCell.FormatCell.bBorderBottom);
                if (border.Bottom.Style != ExcelBorderStyle.None &&
                  moxelCell.FormatCell.BorderColor != Color.Black)
                    border.Bottom.Color.SetColor(moxelCell.FormatCell.BorderColor);
            }

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderLeft))
            {
                border.Left.Style = ExcelWriter.GetBorderStyle(moxelCell.FormatCell.bBorderLeft);
                if (border.Left.Style != ExcelBorderStyle.None &&
                  moxelCell.FormatCell.BorderColor != Color.Black)
                    border.Left.Color.SetColor(moxelCell.FormatCell.BorderColor);
            }

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderTop))
            {
                border.Top.Style = ExcelWriter.GetBorderStyle(moxelCell.FormatCell.bBorderTop);
                if (border.Top.Style != ExcelBorderStyle.None &&
                  moxelCell.FormatCell.BorderColor != Color.Black)
                    border.Top.Color.SetColor(moxelCell.FormatCell.BorderColor);
            }

            if (moxelCell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.BorderRight))
            {
                border.Right.Style = ExcelWriter.GetBorderStyle(moxelCell.FormatCell.bBorderRight);
                if (border.Right.Style != ExcelBorderStyle.None &&
                  moxelCell.FormatCell.BorderColor != Color.Black)
                    border.Right.Color.SetColor(moxelCell.FormatCell.BorderColor);
            }
        }
    }

    public static class ExcelWriter
    {
        public static event ConverterProgressor OnProgress;

        /// <summary>Параметры страницы 1С (ПараметрыСтраницы), заполняются снаружи.</summary>
        public static PageSettings PageSettings = null;

        public static int MaxDegreeOfParallelism = Environment.ProcessorCount / 4 * 3;

        // ---------------- пересчёт единиц (эмпирические коэффициенты, не менять) ----------------

        private const float
            Standard1CSymbol =
                105f; //Стандартный символ 1С = 105 твипов. Это ширина символа "0" шрифтом Arial, размера 10

        private static float
            _standardExcelSymbol = 6f; // Стандартный символ Экселя в пикселях. Шрифт по умолчанию - Calibri, размера 11

        private static readonly float UnitsPerPixel = WinApi.GetUnitsPerPixel();

        private const double Defcolumnwidth = 40.0d;

        private static double MoxelWidthToExcel(double pt)
        {
            var pixels = MoxelWidthToPixels(pt);

            if (pixels <= _standardExcelSymbol)
                return pixels / _standardExcelSymbol;
            else
                return 1 + (pixels - _standardExcelSymbol) / Standard1CSymbol * UnitsPerPixel;
        }

        private static double MoxelWidthToPixels(double pt)
        {
            return pt / 8.23d * Standard1CSymbol / UnitsPerPixel;
        }

        private static double MoxelHeightToExcel(double pt)
        {
            return pt / 4;
        }

        private static float MeasureSymbol(Font font)
        {
            double s1, s2;
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
            {
                s1 = g.MeasureString("0", font).Width;
                s2 = g.MeasureString("00", font).Width;
            }
            var y = 2 * s1 - s2; //Ширина интервала между символами
            var x = s1 - 2 * y; //Ширина символа
            var z = x + y; //Ширина занимаемого места символом в составе строки
            return (float)z;
        }

        private static Size MeasureString(string text, Font font, int areaWidth)
        {
            var constr = new Size { Width = areaWidth, Height = 0 };
            return System.Windows.Forms.TextRenderer.MeasureText(text, font, constr,
                System.Windows.Forms.TextFormatFlags.WordBreak | System.Windows.Forms.TextFormatFlags.VerticalCenter |
                System.Windows.Forms.TextFormatFlags.ExpandTabs);
        }

        internal static ExcelBorderStyle GetBorderStyle(global::Moxel.Moxel.BorderStyle moxelBorder)
        {
            return moxelBorder switch
            {
                BorderStyle.None => ExcelBorderStyle.None,
                BorderStyle.ThinSolid => ExcelBorderStyle.Thin,
                BorderStyle.ThinDotted => ExcelBorderStyle.Dotted,
                BorderStyle.ThinGrayDotted => ExcelBorderStyle.Dotted,
                BorderStyle.ThinDashedShort => ExcelBorderStyle.Dashed,
                BorderStyle.ThinDashedLong => ExcelBorderStyle.DashDotDot,
                BorderStyle.MediumSolid => ExcelBorderStyle.Medium,
                BorderStyle.MediumDashed => ExcelBorderStyle.MediumDashed,
                BorderStyle.ThickSolid => ExcelBorderStyle.Thick,
                BorderStyle.Double => ExcelBorderStyle.Double,
                _ => ExcelBorderStyle.Thin
            };
        }

        // ================================ SAVE ================================

        public static bool Save(global::Moxel.Moxel moxel, string filename)
        {
            if (moxel == null) throw new ArgumentNullException("moxel");
            if (string.IsNullOrEmpty(filename)) throw new ArgumentException("Не задано имя файла.", "filename");

            _standardExcelSymbol = MeasureSymbol(SystemFonts.DefaultFont);

            if (File.Exists(filename))
                File.Delete(filename);

            // --- шрифт по умолчанию (в 1С обычно Arial 8) ---
            var defFontSize = 8;
            var defFontName = "Arial";

            if (moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.FontSize))
            {
                var fs = Math.Abs(moxel.DefFormat.wFontSize) / 4;   // wFontSize знаковый: -4 * размер_пт
                if (fs >= 1) defFontSize = fs;
            }

            if (moxel.FontList != null && moxel.FontList.Count > 0)
            {
                // основной шрифт таблицы — с минимальным ключом списка шрифтов
                var firstKey = moxel.FontList.Keys.Min();
                var faceName = moxel.FontList[firstKey].lfFaceName;
                if (!string.IsNullOrEmpty(faceName))
                    defFontName = faceName;
            }

            var separator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var ms = new MemoryStream())
            using (var package = new ExcelPackage(ms))
            {
                using (var worksheet = package.Workbook.Worksheets.Add("Лист1"))
                {

                    // ---------------- ширины колонок (колонка → умолчание → 40) ----------------
                    for (var columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
                    {
                        double columnwidth = -1;

                        DataCell col;
                        if (moxel.Columns.TryGetValue(columnNumber, out col) &&
                            col.FormatCell.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                            columnwidth = (double)col.FormatCell.wWidth;

                        if (columnwidth < 0 && moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.ColumnWidth))
                            columnwidth = (double)moxel.DefFormat.wWidth;

                        if (columnwidth < 0)
                            columnwidth = Defcolumnwidth;

                        worksheet.Column(columnNumber + 1).Width = MoxelWidthToExcel(columnwidth);
                    }

                    // ---------------- объединения из модели (Unions) ----------------
                    var merges = new List<Mer>();
                    if (moxel.Unions != null)
                        foreach (var union in moxel.Unions)
                        {
                            int c1 = Math.Min(union.dwLeft, union.dwRight), c2 = Math.Max(union.dwLeft, union.dwRight);
                            int r1 = Math.Min(union.dwTop, union.dwBottom), r2 = Math.Max(union.dwTop, union.dwBottom);

                            if (r1 < 0 || c1 < 0) continue;
                            if (r1 == r2 && c1 == c2) continue;                                  // вырожденное
                            if (r2 >= ExcelPackage.MaxRows || c2 >= ExcelPackage.MaxColumns) continue;
                            if (r2 - r1 >= ExcelPackage.MaxRows || c2 - c1 >= ExcelPackage.MaxColumns) continue;

                            var m = new Mer(r1, c1, r2, c2);
                            if (merges.Any(x => x.Intersects(m))) continue;                      // пересечение — битые данные
                            merges.Add(m);
                        }

                    // ================= ФАЗА 1: параллельный расчёт (чистые вычисления) =================
                    // Сюда не входит EPPlus — только модель + GDI-измерения. Результаты
                    // складываются в result[rowNumber] (слоты не пересекаются — синхронизация не нужна).
                    var total = Math.Max(0, moxel.nAllRowCount);
                    var result = new RowOut[total];

                    // Эффективный формат «умолчание листа + колонка» — общий для всех строк
                    var colFormats = new CSheetFormat[Math.Max(0, moxel.nAllColumnCount)];
                    for (var c = 0; c < colFormats.Length; c++)
                    {
                        var f = moxel.DefFormat;
                        DataCell col;
                        if (moxel.Columns.TryGetValue(c, out col))
                            MergeFormat(ref f, col.FormatCell);
                        colFormats[c] = f;
                    }

                    var options = new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, MaxDegreeOfParallelism) };
                    var fontCache = new ConcurrentDictionary<FontKey, Font>();
                    var progressHandler = OnProgress;
                    var done = 0;
                    var lastProgress = 0;

                    if (total > 0)
                        Parallel.For(0, total, options,
                            () => Graphics.FromHwnd(IntPtr.Zero),                 // per-thread Graphics
                            (int rowNumber, ParallelLoopState state, Graphics g) =>
                            {
                                result[rowNumber] = ProcessRow(moxel, rowNumber, colFormats, merges,
                                    fontCache, defFontName, defFontSize, g, separator);

                                if (progressHandler != null)
                                {
                                    var d = Interlocked.Increment(ref done);
                                    var p = (int)((long)d * 100 / total);
                                    var prev = Volatile.Read(ref lastProgress);
                                    if (p - prev > 3 && Interlocked.CompareExchange(ref lastProgress, p, prev) == prev)
                                        progressHandler(p);
                                }
                                return g;
                            },
                            g => g.Dispose());

                    foreach (var fnt in fontCache.Values)
                        fnt.Dispose();

                    // ================= ФАЗА 2: однопоточная запись в EPPlus =================

                    // итоговый список объединений: модель + авто-объединения «по выделению»
                    for (var r = 0; r < total; r++)
                    {
                        var ro = result[r];
                        if (ro == null || ro.AutoMerges == null) continue;
                        foreach (var am in ro.AutoMerges)
                            merges.Add(am);
                    }

                    foreach (var mg in merges)
                        worksheet.Cells[mg.R1 + 1, mg.C1 + 1, mg.R2 + 1, mg.C2 + 1].Merge = true;

                    for (var r = 0; r < total; r++)
                    {
                        var ro = result[r];
                        if (ro == null) continue;

                        if (ro.Cells != null)
                            foreach (var co in ro.Cells)
                                ApplyCell(worksheet, co);

                        worksheet.Row(r + 1).Height = MoxelHeightToExcel(ro.Height);
                    }

                    #region Добавим картинки

                    if (moxel.Objects != null)
                        foreach (var obj in moxel.Objects)
                        {
                            if (obj == null) continue;

                            switch (obj.Picture.dwType)
                            {
                                case ObjectType.Ole:
                                case ObjectType.Picture:
                                    {
                                        // В memory-режиме v1 полезная нагрузка картинок не извлекается
                                        // (pObject == null) — экспортировать заглушку 1x1 не нужно.
                                        if (obj.pObject == null) 
                                            break;

                                        using (var pms = new MemoryStream())
                                        {
                                            //Странныый косяк с GDI+. Без такого финта выдает неопознанную ошибку
                                            using (Bitmap bmp = obj.pObject)
                                            {
                                                bmp.Save(pms, ImageFormat.Png);

                                                // итоговый размер в пикселях = размер области объекта;
                                                // при нулевой/кривой области не делим на ноль, берём натуральный размер
                                                int targetW, targetH;
                                                var area = obj.ImageArea;
                                                if (area.Width > 0 && area.Height > 0)
                                                {
                                                    var zoomfactorX = (double)bmp.Width / area.Width;
                                                    var zoomfactorY = (double)bmp.Height / area.Height;
                                                    targetW = (int)Math.Round(bmp.Width / zoomfactorX);
                                                    targetH = (int)Math.Round(bmp.Height / zoomfactorY);
                                                }
                                                else
                                                {
                                                    targetW = bmp.Width;
                                                    targetH = bmp.Height;
                                                }

                                                using (var picture = worksheet.Drawings.AddPicture($"D{obj.Picture.dwZOrder}", pms))
                                                {
                                                    picture.SetPosition(Math.Max(0, obj.Picture.dwRowStart), Math.Max(0, obj.Picture.dwOffsetTop / 3),
                                                                        Math.Max(0, obj.Picture.dwColumnStart), Math.Max(0, obj.Picture.dwOffsetLeft / 3));
                                                    picture.SetSize(Math.Max(1, targetW), Math.Max(1, targetH));
                                                }
                                            }
                                        }
                                        break;
                                    }
                                case ObjectType.Text:
                                    {
                                        using (var textBox = worksheet.Drawings.AddShape($"D{obj.Picture.dwZOrder}", eShapeStyle.Rect))
                                        {
                                            MoxelRow row;
                                            moxel.Rows.TryGetValue(obj.Picture.dwRowStart, out row);

                                            if (!string.IsNullOrEmpty(obj.Text))
                                            {
                                                var text = textBox.RichText.Add(obj.Text);
                                                text.SetFromFont(obj.GetFont(defFontName, defFontSize));
                                            }

                                            if (obj.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontColor))
                                            {
                                                textBox.Font.Fill.Color = obj.FormatCell.FontColor;
                                            }
                                            else
                                            {
                                                textBox.Font.Fill.Color = Color.Black;
                                            }

                                            var area = obj.ImageArea;
                                            if (area.Width > 0 && area.Height > 0)
                                                textBox.SetSize(area.Width, area.Height);

                                            if (obj.FormatCell.dwFlags.HasFlag(MoxelCellFlags.Background))
                                            {
                                                textBox.Fill.Color = obj.FormatCell.BgColor;
                                            }

                                            textBox.SetPosition(Math.Max(0, obj.Picture.dwRowStart), Math.Max(0, obj.Picture.dwOffsetTop / 3),
                                                                Math.Max(0, obj.Picture.dwColumnStart), Math.Max(0, obj.Picture.dwOffsetLeft / 3));
                                        }
                                        break;
                                    }
                                case ObjectType.Line:
                                    break;
                                case ObjectType.Rectangle:
                                    break;
                                case ObjectType.None:
                                    break;
                                default:
                                    break;
                            }
                        }

                    #endregion

                    #region Колонтитулы
                    // 1С-коды: #D дата, #T время, #P страница, #Q всего страниц.
                    // Пишем в Odd*: при differentOddEven=false (умолчание) в Excel видны именно они.
                    if (moxel.Header != null && !string.IsNullOrEmpty(moxel.Header.Text))
                    {
                        var header = MapHeaderFooterCodes(moxel.Header.Text);
                        worksheet.HeaderFooter.OddHeader.LeftAlignedText = header;
                        worksheet.HeaderFooter.EvenHeader.LeftAlignedText = header;
                    }

                    if (moxel.Footer != null && !string.IsNullOrEmpty(moxel.Footer.Text))
                    {
                        var footer = MapHeaderFooterCodes(moxel.Footer.Text);
                        worksheet.HeaderFooter.OddFooter.LeftAlignedText = footer;
                        worksheet.HeaderFooter.EvenFooter.LeftAlignedText = footer;
                    }
                    #endregion

                    #region Разрывы страниц

                    if (moxel.HorisontalPageBreaks != null)
                        foreach (int br in moxel.HorisontalPageBreaks)
                            if (br >= 0 && br < ExcelPackage.MaxRows - 1)
                                worksheet.Row(br + 1).PageBreak = true;

                    if (moxel.VerticalPageBreaks != null)
                        foreach (int br in moxel.VerticalPageBreaks)
                            if (br >= 0 && br < ExcelPackage.MaxColumns - 1)
                                worksheet.Column(br + 1).PageBreak = true;

                    #endregion

                    #region Параметры страницы

                    if (PageSettings != null)
                    {
                        // Ориентация: 1С — 1 портрет / 2 ландшафт. Маппим по значению,
                        // не полагаясь на нумерацию enum eOrientation.
                        var orient = (int)PageSettings.Get(PageSettings.OptionType.Orient);
                        worksheet.PrinterSettings.Orientation = orient == 2 ? eOrientation.Landscape : eOrientation.Portrait;

                        var paper = (int)PageSettings.Get(PageSettings.OptionType.Paper);
                        if (paper > 0)
                            worksheet.PrinterSettings.PaperSize = (ePaperSize)paper;

                        if ((int)PageSettings.Get(PageSettings.OptionType.FitToPage) == 1)
                            worksheet.PrinterSettings.FitToPage = true;
                        else
                        {
                            var scale = (int)PageSettings.Get(PageSettings.OptionType.Scale);
                            if (scale > 0)
                                worksheet.PrinterSettings.Scale = scale;
                        }

                        worksheet.PrinterSettings.BlackAndWhite = (int)PageSettings.Get(PageSettings.OptionType.BlackAndWhite) == 1;

                        // поля 1С — в мм, в Excel — в дюймах
                        worksheet.PrinterSettings.BottomMargin = ((int)PageSettings.Get(PageSettings.OptionType.Bottom)) / 25.4m;
                        worksheet.PrinterSettings.LeftMargin = ((int)PageSettings.Get(PageSettings.OptionType.Left)) / 25.4m;
                        worksheet.PrinterSettings.RightMargin = ((int)PageSettings.Get(PageSettings.OptionType.Right)) / 25.4m;
                        worksheet.PrinterSettings.TopMargin = ((int)PageSettings.Get(PageSettings.OptionType.Top)) / 25.4m;

                        worksheet.PrinterSettings.FooterMargin =
                            ((int)PageSettings.Get(PageSettings.OptionType.Footer)) / 25.4m;
                        worksheet.PrinterSettings.HeaderMargin =
                            ((int)PageSettings.Get(PageSettings.OptionType.Header)) / 25.4m;

                        var repeatRowFrom = (int)PageSettings.Get(PageSettings.OptionType.RepeatRowFrom);
                        if (repeatRowFrom != 0)
                        {
                            var repeatRowTo = (int)PageSettings.Get(PageSettings.OptionType.RepeatRowTo);
                            if (repeatRowTo < repeatRowFrom) repeatRowTo = repeatRowFrom;
                            worksheet.PrinterSettings.RepeatRows =
                                new ExcelAddress(repeatRowFrom + 1, 1, repeatRowTo + 1, ExcelPackage.MaxColumns);
                        }

                        var repeatColFrom = (int)PageSettings.Get(PageSettings.OptionType.RepeatColFrom);
                        if (repeatColFrom != 0)
                        {
                            var repeatColTo = (int)PageSettings.Get(PageSettings.OptionType.RepeatColTo);
                            if (repeatColTo < repeatColFrom) repeatColTo = repeatColFrom;
                            worksheet.PrinterSettings.RepeatColumns =
                                new ExcelAddress(1, repeatColFrom + 1, ExcelPackage.MaxRows, repeatColTo + 1);
                        }

                        if ((int)PageSettings.Get(PageSettings.OptionType.RangeBottom) != 0)
                        {
                            var rangeTop = (int)PageSettings.Get(PageSettings.OptionType.RangeTop);
                            var rangeLeft = (int)PageSettings.Get(PageSettings.OptionType.RangeLeft);
                            var rangeBottom = (int)PageSettings.Get(PageSettings.OptionType.RangeBottom);
                            var rangeRight = (int)PageSettings.Get(PageSettings.OptionType.RangeRight);

                            if (rangeTop >= 0 && rangeLeft >= 0 && rangeBottom >= rangeTop && rangeRight >= rangeLeft)
                            {
                                var printArea = worksheet.Cells[
                                    rangeTop + 1, rangeLeft + 1, rangeBottom + 1, rangeRight + 1];
                                worksheet.PrinterSettings.PrintArea = printArea;
                            }
                        }
                    }
                    #endregion

                    package.Save();
                }

                File.WriteAllBytes(filename, ms.ToArray());
            }

            return File.Exists(filename);
        }

        // ============================ ФАЗА 1 ============================

        /// <summary>Расчёт одной строки таблицы (чистая функция, потокобезопасна).</summary>
        private static RowOut ProcessRow(Moxel moxel, int rowNumber, CSheetFormat[] colFormats,
            List<Mer> unions, ConcurrentDictionary<FontKey, Font> fontCache,
            string defFontName, int defFontSize, Graphics g, char separator)
        {
            var ro = new RowOut();

            MoxelRow row;
            moxel.Rows.TryGetValue(rowNumber, out row);

            // ---------------- высота строки ----------------
            double rowHeight;
            bool autoHeight;

            if (row != null)
            {
                var h = row.Height;                 // моксел-единицы; 0 = авторасчёт
                if (h == 0) { autoHeight = true; rowHeight = 0; }
                else { autoHeight = false; rowHeight = h; }
            }
            else
            {
                autoHeight = false;
                rowHeight = moxel.DefFormat.dwFlags.HasFlag(MoxelCellFlags.RowHeight)
                    ? moxel.DefFormat.wHeight
                    : 45;
                if (rowHeight <= 0) rowHeight = 45;
            }

            // ---------------- авто-объединения «по выделению» ----------------
            // Ячейка с текстом и переносом, у которой нет правой границы, а сосед справа
            // пуст и помечен выравниванием BySelection — продолжение объединения.
            var autoMerges = new List<Mer>();
            if (row != null && row.values != null)
                foreach (var kv in row.values)
                {
                    var rawText = kv.Value.Text;
                    if (string.IsNullOrEmpty(rawText)) continue;

                    var c = kv.Key;
                    if (c < 0 || c + 1 >= moxel.nAllColumnCount) continue;

                    var eff = GetEffectiveFormat(colFormats[c], row, kv.Value);

                    if (eff.bControlContent != TextControl.Wrap ||
                        (eff.dwFlags & MoxelCellFlags.BorderRight) != 0) continue;

                    DataCell next;
                    if (row.TryGetValue(c + 1, out next) && !string.IsNullOrEmpty(next.Text)) continue;

                    var cn = c + 1;
                    while (cn < moxel.nAllColumnCount)          // граница цикла обязательна
                    {
                        DataCell nc;
                        var nf = row.TryGetValue(cn, out nc)
                            ? GetEffectiveFormat(colFormats[cn], row, nc)
                            : GetEffectiveFormat(colFormats[cn], row, null);

                        if (nf.bHorAlign.HasFlag(TextHorzAlign.BySelection) &&
                            nf.bBorderRight == BorderStyle.None)
                            cn++;
                        else
                            break;
                    }

                    if (cn > c + 1)
                    {
                        var m = new Mer(rowNumber, c, rowNumber, cn - 1);
                        // не пересекаем ни объединения модели, ни уже найденные авто-объединения
                        if (!unions.Any(u => u.Intersects(m)) &&
                            !autoMerges.Any(a => a.Intersects(m)))
                            autoMerges.Add(m);
                    }
                }
            ro.AutoMerges = autoMerges;

            // ---------------- ячейки строки ----------------
            var cells = new List<CellOut>();
            ro.Cells = cells;

            var measuredHeight = 0d;

            for (var columnNumber = 0; columnNumber < moxel.nAllColumnCount; columnNumber++)
            {
                DataCell moxelCell = null;
                var hasCell = row != null && row.TryGetValue(columnNumber, out moxelCell);
                var eff = GetEffectiveFormat(colFormats[columnNumber], row, moxelCell);

                // объединение?
                Mer merge;
                bool topLeft;
                FindMerge(unions, autoMerges, rowNumber, columnNumber, out merge, out topLeft);

                // содержимое
                string rawText = null;
                string text = null;
                double number = 0;
                bool isNumber = false;
                string numberFormat = null;

                if (moxelCell != null && !string.IsNullOrEmpty(moxelCell.Text))
                {
                    rawText = moxelCell.Text;
                    if (TryGetNumber(moxelCell, rawText, separator, out number, out numberFormat))
                    {
                        isNumber = true;
                        text = rawText;
                    }
                    else
                    {
                        text = rawText.TrimEnd('\r', '\n');
                    }
                }

                var hasText = isNumber || !string.IsNullOrEmpty(text);

                // рамки/заливка — есть и у ячеек без текста (унаследованные форматы)
                var flags = eff.dwFlags;
                var hasBorder = HasAnyBorder(eff);
                var hasFill = (flags & MoxelCellFlags.Background) != 0 &&
                              eff.BgColor != Color.White && eff.BgColor != Color.Empty;

                if (!hasText && !hasBorder && !hasFill)
                    continue;                                                    // пустая ячейка без стилей

                if (merge.R1 >= 0 && !topLeft && !hasBorder)
                    continue;                                                    // внутри объединения: рамки держим, контент — только в левой верхней

                var co = new CellOut
                {
                    Row = rowNumber,
                    Col = columnNumber,
                    Eff = eff,
                    Merge = merge,
                    TopLeft = topLeft,
                    HasText = hasText && (merge.R1 < 0 || topLeft),
                    IsNumber = isNumber,
                    Number = number,
                    NumberFormat = numberFormat,
                    Text = isNumber ? null : text
                };

                if (co.HasText)
                {
                    // ---- шрифт (эффективный) ----
                    var fname = defFontName;
                    if ((flags & MoxelCellFlags.FontName) != 0)
                    {
                        LOGFONT lf;
                        if (moxel.FontList != null &&
                            moxel.FontList.TryGetValue(eff.wFontNumber, out lf) &&
                            !string.IsNullOrEmpty(lf.lfFaceName))
                            fname = lf.lfFaceName;
                    }

                    var fsize = (float)defFontSize;
                    if ((flags & MoxelCellFlags.FontSize) != 0)
                    {
                        var s = Math.Abs(eff.wFontSize) / 4f;   // знаковый: -4*пт
                        if (s >= 1f) fsize = s;
                    }

                    var bold = (flags & MoxelCellFlags.FontWeight) != 0 && eff.bFontBold == clFontWeight.Bold;
                    var italic = (flags & MoxelCellFlags.FontItalic) != 0 && eff.bFontItalic;
                    var underline = (flags & MoxelCellFlags.FontUnderline) != 0 && eff.bFontUnderline;

                    co.FontName = fname;
                    co.FontSize = fsize;
                    co.Bold = bold;
                    co.Italic = italic;
                    co.Underline = underline;

                    // ---- перенос по словам / контроль отображения ----
                    var wrap = rawText.Contains("\r\n");
                    if ((flags & MoxelCellFlags.Control) != 0)
                    {
                        switch (eff.bControlContent)
                        {
                            case TextControl.Auto:
                                // «Авто»: текст переливается в пустого соседа справа — Excel
                                // этого не умеет, переносим только когда сосед непуст
                                DataCell nb;
                                var nextEmpty = !(row.TryGetValue(columnNumber + 1, out nb)) ||
                                                string.IsNullOrEmpty(nb.Text);
                                wrap = !nextEmpty;
                                break;
                            case TextControl.Wrap:
                                wrap = true;
                                break;
                            case TextControl.Cut:
                                wrap = false;
                                break;
                            // Fill / Red / FillAndRed — на перенос не влияют
                        }
                    }
                    co.Wrap = wrap;

                    co.Rotation = MapTextRotation(moxelCell.TextOrientation);

                    // ---- авторасчёт высоты строки ----
                    if (autoHeight)
                    {
                        var rowsSpanned = merge.R1 >= 0 ? merge.R2 - merge.R1 + 1 : 1;
                        var colSpan = merge.R1 >= 0 ? merge.C2 - merge.C1 + 1 : 1;
                        var areaWidth = MoxelWidthToPixels(moxel.GetWidth(columnNumber,
                            Math.Min(columnNumber + colSpan, moxel.nAllColumnCount)));

                        var measureFont = fontCache.GetOrAdd(new FontKey(fname, fsize, bold),
                            k => new Font(k.Name, k.Size, k.Bold ? FontStyle.Bold : FontStyle.Regular));

                        SizeF stringSize;
                        if (!wrap)
                            stringSize = g.MeasureString(rawText, measureFont);
                        else
                            stringSize = MeasureString(rawText, measureFont, (int)Math.Round(areaWidth));

                        var heigth = (int)Math.Ceiling(stringSize.Height / rowsSpanned / 1.27 * 4);
                        if (heigth > measuredHeight) measuredHeight = heigth;

                        // паритет с прежней версией: у авторасчётных строк Bottom+перенос
                        // без поворота превращается в Justify
                        if ((flags & MoxelCellFlags.AlignV) != 0 &&
                            eff.bVertAlign == TextVertAlign.Bottom && wrap && co.Rotation == 0)
                            co.AutoJustifyV = true;
                    }
                }

                cells.Add(co);
            }

            ro.Height = autoHeight ? (measuredHeight > 0 ? Math.Max(measuredHeight, 45) : 45) : rowHeight;
            return ro;
        }

        // ============================ ФАЗА 2 ============================

        /// <summary>Перенос одной ячейки в EPPlus (однопоточно).</summary>
        private static void ApplyCell(ExcelWorksheet ws, CellOut co)
        {
            var merged = co.Merge.R1 >= 0 && co.TopLeft;
            ExcelRange range = merged
                ? ws.Cells[co.Merge.R1 + 1, co.Merge.C1 + 1, co.Merge.R2 + 1, co.Merge.C2 + 1]
                : ws.Cells[co.Row + 1, co.Col + 1];

            if (co.HasText)
            {
                if (co.IsNumber)
                {
                    range.Value = co.Number;
                    if (!string.IsNullOrEmpty(co.NumberFormat))
                        range.Style.Numberformat.Format = co.NumberFormat;
                }
                else
                {
                    range.Value = co.Text;
                }

                var style = range.Style;

                style.Font.SetFromFont(co.FontName, co.FontSize, co.Bold, co.Italic, co.Underline);

                var flags = co.Eff.dwFlags;
                if ((flags & MoxelCellFlags.FontColor) != 0)
                    style.Font.Color.SetColor(co.Eff.FontColor);

                style.WrapText = co.Wrap;

                if (co.Rotation != 0)
                    style.TextRotation = co.Rotation;

                if ((flags & MoxelCellFlags.AlignH) != 0)
                    MapHorzAlign(co.Eff.bHorAlign, style);

                if ((flags & MoxelCellFlags.AlignV) != 0)
                    MapVertAlign(co.Eff.bVertAlign, style);

                if (co.AutoJustifyV)
                    style.VerticalAlignment = ExcelVerticalAlignment.Justify;
            }

            // рамки — с эффективного формата (в т.ч. унаследованные от колонки/строки)
            var m = co.Eff.dwFlags;
            var border = range.Style.Border;

            if ((m & MoxelCellFlags.BorderLeft) != 0 && co.Eff.bBorderLeft != BorderStyle.None)
            {
                border.Left.Style = GetBorderStyle(co.Eff.bBorderLeft);
                if (co.Eff.BorderColor != Color.Black && co.Eff.BorderColor != Color.Empty)
                    border.Left.Color.SetColor(co.Eff.BorderColor);
            }

            if ((m & MoxelCellFlags.BorderTop) != 0 && co.Eff.bBorderTop != BorderStyle.None)
            {
                border.Top.Style = GetBorderStyle(co.Eff.bBorderTop);
                if (co.Eff.BorderColor != Color.Black && co.Eff.BorderColor != Color.Empty)
                    border.Top.Color.SetColor(co.Eff.BorderColor);
            }

            if ((m & MoxelCellFlags.BorderRight) != 0 && co.Eff.bBorderRight != BorderStyle.None)
            {
                border.Right.Style = GetBorderStyle(co.Eff.bBorderRight);
                if (co.Eff.BorderColor != Color.Black && co.Eff.BorderColor != Color.Empty)
                    border.Right.Color.SetColor(co.Eff.BorderColor);
            }

            if ((m & MoxelCellFlags.BorderBottom) != 0 && co.Eff.bBorderBottom != BorderStyle.None)
            {
                border.Bottom.Style = GetBorderStyle(co.Eff.bBorderBottom);
                if (co.Eff.BorderColor != Color.Black && co.Eff.BorderColor != Color.Empty)
                    border.Bottom.Color.SetColor(co.Eff.BorderColor);
            }

            // заливка (исправлено: раньше проверялся BorderColor вместо BgColor)
            if ((m & MoxelCellFlags.Background) != 0 &&
                co.Eff.BgColor != Color.White && co.Eff.BgColor != Color.Empty)
            {
                var fill = range.Style.Fill;
                fill.PatternType = ExcelFillStyle.Solid;
                fill.BackgroundColor.SetColor(co.Eff.BgColor);
            }
        }

        // ============================ СЛУЖЕБНОЕ ============================

        /// <summary>
        /// Наложение слоя формата на результат (по битам маски).
        /// dst — накопленный формат, src — очередной слой (колонка / строка / ячейка).
        /// </summary>
        private static void MergeFormat(ref CSheetFormat dst, CSheetFormat src)
        {
            var m = src.dwFlags;
            if (m == 0) return;

            if ((m & MoxelCellFlags.FontName) != 0) { dst.wFontNumber = src.wFontNumber; dst.dwFlags |= MoxelCellFlags.FontName; }
            if ((m & MoxelCellFlags.FontSize) != 0) { dst.wFontSize = src.wFontSize; dst.dwFlags |= MoxelCellFlags.FontSize; }
            if ((m & MoxelCellFlags.FontWeight) != 0) { dst.bFontBold = src.bFontBold; dst.dwFlags |= MoxelCellFlags.FontWeight; }
            if ((m & MoxelCellFlags.FontItalic) != 0) { dst.bFontItalic = src.bFontItalic; dst.dwFlags |= MoxelCellFlags.FontItalic; }
            if ((m & MoxelCellFlags.FontUnderline) != 0) { dst.bFontUnderline = src.bFontUnderline; dst.dwFlags |= MoxelCellFlags.FontUnderline; }
            if ((m & MoxelCellFlags.BorderLeft) != 0) { dst.bBorderLeft = src.bBorderLeft; dst.dwFlags |= MoxelCellFlags.BorderLeft; }
            if ((m & MoxelCellFlags.BorderTop) != 0) { dst.bBorderTop = src.bBorderTop; dst.dwFlags |= MoxelCellFlags.BorderTop; }
            if ((m & MoxelCellFlags.BorderRight) != 0) { dst.bBorderRight = src.bBorderRight; dst.dwFlags |= MoxelCellFlags.BorderRight; }
            if ((m & MoxelCellFlags.BorderBottom) != 0) { dst.bBorderBottom = src.bBorderBottom; dst.dwFlags |= MoxelCellFlags.BorderBottom; }
            if ((m & MoxelCellFlags.BorderColor) != 0) { dst.bBorderColor = src.bBorderColor; dst.dwFlags |= MoxelCellFlags.BorderColor; }
            if ((m & MoxelCellFlags.RowHeight) != 0) { dst.wHeight = src.wHeight; dst.dwFlags |= MoxelCellFlags.RowHeight; }
            if ((m & MoxelCellFlags.ColumnWidth) != 0) { dst.wWidth = src.wWidth; dst.dwFlags |= MoxelCellFlags.ColumnWidth; }
            if ((m & MoxelCellFlags.AlignH) != 0) { dst.bHorAlign = src.bHorAlign; dst.dwFlags |= MoxelCellFlags.AlignH; }
            if ((m & MoxelCellFlags.AlignV) != 0) { dst.bVertAlign = src.bVertAlign; dst.dwFlags |= MoxelCellFlags.AlignV; }
            if ((m & MoxelCellFlags.FontColor) != 0) { dst.bFontColor = src.bFontColor; dst.dwFlags |= MoxelCellFlags.FontColor; }
            if ((m & MoxelCellFlags.Background) != 0) { dst.bBackground = src.bBackground; dst.dwFlags |= MoxelCellFlags.Background; }
            if ((m & MoxelCellFlags.PatternType) != 0) { dst.bPatternType = src.bPatternType; dst.dwFlags |= MoxelCellFlags.PatternType; }
            if ((m & MoxelCellFlags.PatternColor) != 0) { dst.bPatternColor = src.bPatternColor; dst.dwFlags |= MoxelCellFlags.PatternColor; }
            if ((m & MoxelCellFlags.Control) != 0) { dst.bControlContent = src.bControlContent; dst.dwFlags |= MoxelCellFlags.Control; }
            if ((m & MoxelCellFlags.Type) != 0) { dst.bType = src.bType; dst.dwFlags |= MoxelCellFlags.Type; }
            if ((m & MoxelCellFlags.Protect) != 0) { dst.bAllowEdit = src.bAllowEdit; dst.dwFlags |= MoxelCellFlags.Protect; }
        }

        /// <summary>Эффективный формат: «умолчание+колонка» (baseAndColumn) ← строка ← ячейка.</summary>
        private static CSheetFormat GetEffectiveFormat(CSheetFormat baseAndColumn, MoxelRow row, DataCell cell)
        {
            var f = baseAndColumn;              // копия структуры
            if (row != null) MergeFormat(ref f, row.FormatCell);
            if (cell != null) MergeFormat(ref f, cell.FormatCell);
            return f;
        }

        private static bool HasAnyBorder(CSheetFormat f)
        {
            var m = f.dwFlags;
            return
                ((m & MoxelCellFlags.BorderLeft) != 0 && f.bBorderLeft != BorderStyle.None) ||
                ((m & MoxelCellFlags.BorderTop) != 0 && f.bBorderTop != BorderStyle.None) ||
                ((m & MoxelCellFlags.BorderRight) != 0 && f.bBorderRight != BorderStyle.None) ||
                ((m & MoxelCellFlags.BorderBottom) != 0 && f.bBorderBottom != BorderStyle.None);
        }

        /// <summary>Поиск объединения, содержащего ячейку (сначала модельные, потом авто).</summary>
        private static void FindMerge(List<Mer> unions, List<Mer> autoMerges, int row, int col, out Mer merge, out bool topLeft)
        {
            for (var i = 0; i < unions.Count; i++)
                if (unions[i].Contains(row, col))
                {
                    merge = unions[i];
                    topLeft = merge.R1 == row && merge.C1 == col;
                    return;
                }

            for (var i = 0; i < autoMerges.Count; i++)
                if (autoMerges[i].Contains(row, col))
                {
                    merge = autoMerges[i];
                    topLeft = merge.R1 == row && merge.C1 == col;
                    return;
                }

            merge = default(Mer);
            merge.R1 = -1;
            topLeft = false;
        }

        /// <summary>
        /// Числовое значение ячейки. Сначала родное значение (бит Value — результат
        /// выражения 1С; из памяти это m_strDetails), затем прежняя эвристика по
        /// точкам/запятым. Родное значение применяется только к «числоподобному»
        /// тексту, чтобы не терять форматированные величины («123 руб.», «10,5%»).
        /// </summary>
        private static bool TryGetNumber(DataCell cell, string text, char separator, out double value, out string format)
        {
            value = 0;
            format = null;

            // 1) родное значение ячейки
            var v = cell.Value;
            if (!string.IsNullOrEmpty(v) && IsNumericLike(text))
            {
                double nv;
                if (double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out nv) ||
                    double.TryParse(v, NumberStyles.Float, CultureInfo.CurrentCulture, out nv))
                {
                    // Excel не умеет Infinity/NaN — такие значения оставляем текстом
                    if (!double.IsInfinity(nv) && !double.IsNaN(nv))
                    {
                        value = nv;
                        format = nv == Math.Floor(nv) && Math.Abs(nv) < 1E15 ? "#,##0" : "#,##0.00";
                        return true;
                    }
                }
            }

            // 2) эвристика по тексту
            int dots = 0, commas = 0;
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                if (ch == '.') dots++;
                else if (ch == ',') commas++;
            }

            if (dots == 0 && commas == 0) return false;

            string t;
            if (dots == 1 && commas == 1)
                t = text.Replace(",", "").Replace('.', separator);
            else if (dots == 0 && commas == 1)
                t = text.Replace(',', separator);
            else if (dots == 1 && commas == 0)
                t = text.Replace('.', separator);
            else
                return false;

            double tv;
            if (!string.IsNullOrEmpty(t) && double.TryParse(t, out tv))
            {
                value = tv;
                format = "#,##0.00";
                return true;
            }

            return false;
        }

        /// <summary>Текст состоит только из цифр/пробелов/разделителей — кандидат в числа.</summary>
        private static bool IsNumericLike(string text)
        {
            var hasDigit = false;
            foreach (var ch in text)
            {
                if (char.IsDigit(ch)) { hasDigit = true; continue; }
                if (ch == ' ' || ch == '\t' || ch == '\u00A0' ||
                    ch == '.' || ch == ',' || ch == '-' || ch == '+') continue;
                return false;
            }
            return hasDigit;
        }

        /// <summary>
        /// Поворот текста: EPPlus принимает 0..180 и 255 (вертикальный стек).
        /// 1С-овские 270° (сверху вниз) в Excel прямого аналога не имеют —
        /// используем 90° (вертикальный текст, читается снизу вверх).
        /// </summary>
        private static int MapTextRotation(short orientation)
        {
            if (orientation <= 0) return 0;
            if (orientation <= 180) return orientation;
            if (orientation == 255) return 255;
            return 90;
        }

        private static void MapHorzAlign(TextHorzAlign align, ExcelStyle style)
        {
            switch (align)
            {
                case TextHorzAlign.Left:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    break;
                case TextHorzAlign.Right:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    break;
                case TextHorzAlign.Center:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    break;
                case TextHorzAlign.Justify:
                    style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                    break;
                // BySelection (0x20) — признак объединения, на выравнивание не влияет
            }
        }

        private static void MapVertAlign(TextVertAlign align, ExcelStyle style)
        {
            switch (align)
            {
                case TextVertAlign.Bottom:
                    style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                    break;
                case TextVertAlign.Middle:
                    style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    break;
                case TextVertAlign.Top:
                    style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    break;
            }
        }

        private static string MapHeaderFooterCodes(string text)
        {
            return text.Replace("#D", "&D").Replace("#T", "&T")
                       .Replace("#P", "&P").Replace("#Q", "&N");
        }

        // ============================ СТРУКТУРЫ ============================

        /// <summary>Прямоугольник объединённых ячеек, 0-based, включительно.</summary>
        private struct Mer
        {
            public int R1, C1, R2, C2;

            public Mer(int r1, int c1, int r2, int c2) { R1 = r1; C1 = c1; R2 = r2; C2 = c2; }

            public bool Contains(int row, int col)
            {
                return row >= R1 && row <= R2 && col >= C1 && col <= C2;
            }

            public bool Intersects(Mer o)
            {
                return R1 <= o.R2 && o.R1 <= R2 && C1 <= o.C2 && o.C1 <= C2;
            }
        }

        /// <summary>Ключ кэша GDI-шрифтов для измерения высот.</summary>
        private struct FontKey : IEquatable<FontKey>
        {
            public readonly string Name;
            public readonly float Size;
            public readonly bool Bold;

            public FontKey(string name, float size, bool bold)
            {
                Name = name;
                Size = size;
                Bold = bold;
            }

            public bool Equals(FontKey other)
            {
                return Bold == other.Bold && Size.Equals(other.Size) && string.Equals(Name, other.Name);
            }

            public override bool Equals(object obj)
            {
                return obj is FontKey && Equals((FontKey)obj);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    var h = (Name != null ? Name.GetHashCode() : 0);
                    h = (h * 397) ^ Size.GetHashCode();
                    h = (h * 397) ^ (Bold ? 1 : 0);
                    return h;
                }
            }
        }

        /// <summary>Результат расчёта одной строки (фаза 1).</summary>
        private sealed class RowOut
        {
            /// <summary>Высота строки в моксел-единицах (уже включён авторасчёт).</summary>
            public double Height;

            public List<CellOut> Cells;

            /// <summary>Авто-объединения «по выделению» этой строки.</summary>
            public List<Mer> AutoMerges;
        }

        /// <summary>Результат расчёта одной ячейки (фаза 1) — переносится в EPPlus (фаза 2).</summary>
        private sealed class CellOut
        {
            public int Row, Col;

            /// <summary>Есть значение (только для левой верхней ячейки объединения / обычных ячеек).</summary>
            public bool HasText;
            public bool IsNumber;
            public double Number;
            public string NumberFormat;
            public string Text;

            public string FontName;
            public float FontSize;
            public bool Bold, Italic, Underline;
            public bool Wrap;
            public int Rotation;
            public bool AutoJustifyV;

            /// <summary>Эффективный формат (умолчание ← колонка ← строка ← ячейка).</summary>
            public CSheetFormat Eff;

            /// <summary>Объединение, содержащее ячейку (R1 &lt; 0 — без объединения).</summary>
            public Mer Merge;
            public bool TopLeft;
        }
    }
}
