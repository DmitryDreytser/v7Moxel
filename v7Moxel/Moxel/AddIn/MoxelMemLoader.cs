// ============================================================================
// MoxelMemLoader.cs — заполнение модели Moxel НАПРЯМУЮ из памяти 1С 7.7.
//
// Замыкает цепочку: V7Table.SetTable (указатель CSheetDoc) → MemSheet
// (безопасный ридер CSheet) → ЗДЕСЬ: готовый Moxel для ExcelWriter /
// HtmlWriter / PDFWriter. На этом пути НЕ нужны: сериализация (ReadMoxel),
// CArchive, vtable-patch CFile::Write, разбор mxl и вообще файловая
// инфраструктура (NativeMethods/OLE32/HRESULT, IStorage/IStream/StgOpenStorage).
//
// Соответствие модели (проверено по MOXEL.H и исходникам модели):
//
//   Moxel.nAllColumnCount/nAllRowCount ← CSheet::m_nCols@0x164 / m_nRows@0x168
//   Moxel.DefFormat                    ← базовый CSheetFormat листа (@0)
//   Moxel.FontList<int,LOGFONT>        ← m_FontsArray@0x84 = CSortArray<int,
//                                           LOGFONT>, LOGFONTA по значению
//                                           (0x3C); ключи — из m_indexes
//   Moxel.Columns<int,DataCell>        ← m_ColumnsArray@0x2C = CSortArray<int,
//                                           CSheetFormat*>
//   Moxel.Rows<int,MoxelRow>           ← m_RowsArray@0x58 = CSortArray<int,
//                                           CSheetRow*>
//   MoxelRow.FormatCell                ← базовый CSheetFormat строки (0x24)
//   MoxelRow.values<int,DataCell>      ← CSheetRow::m_Keys@0x28 / m_Cells@0x3C
//   Moxel.Header / Footer              ← CSheetCell m_Header@0xDC / m_Footer@0x120
//   Moxel.Objects (EmbeddedObject)     ← m_SheetDrawingList@0x1AC CList<CSheetDrawing*>
//       Picture (файл, 40 байт) ↔ хвост CSheetDrawing 0x4C..0x74 — порядок
//       полей совпадает байт-в-байт: m_type, SHEETRECT{TL,BR} из SHEETPOINT
//       {xCell,yCell,xOff,yOff}, m_DrawingIndex
//   Moxel.Unions (CellsUnion)          ← UnkArr2@0x1F4 / UnkArr3@0x208
//       CPtrArray → CRect* {l,t,r,b} = байт-в-байт CellsUnion (гипотеза —
//       калибруется дампом MemSheet.Dump)
//   Moxel.Horisontal/VerticalSections  ← m_HorzSectionArr@0x1CC /
//       m_VertSectionArr@0x1E0, CArray<CSheetOutline>; CSheetOutline
//       {vft,Start,End,m_data,Name} ↔ файловая Section{Begin,End,Level,Name}
//   Moxel.Horisontal/VerticalPageBreaks← два CArray<int> в зоне 0x184..0x1AC;
//       PB1→H, PB2→V (порядок как у секций H-затем-V; сверить на живой 1С)
//   Moxel.AreaNames (MoxelArea)        ← m_SheetNames@0x16C CSheetNames (0x18)
//       = SGI-хэш VC6 (std::hash_map): vector buckets {start@+4, finish@+8,
//       end_of_storage@+C} + num_elements@+0x10; узел {next@0, CString key@4,
//       CSheetNamedItem val@8}; CSheetNamedItem{Type sntCell=1/sntDraw=2,
//       DrawID, CSheetSelection{c1,r1,c2,r2}} ↔ файловая Area{Unknown1=1,
//       Unknown2=DrawID, AreaType=Type, ColumnBegin=c1, RowBegin=r1,
//       ColumnEnd=c2, RowEnd=r2}
//
// ЗАПОЛНЕНИЕ ФОРМАТОВ: DataCell.FormatCell = СОБСТВЕННЫЙ формат объекта
// (не эффективный): dwFlags = сырая маска объекта, байты полей копируются
// ВСЕ (ровно как пишет родной CSheetFormat::Serialize — блок 30+2 байта
// без vftable). Наслоение default←колонка←строка←ячейка выполняет модель
// при рендере (GetColumnWidth/GetRowHeight/MoxelRow-индексатор). Биты маски
// памяти == MoxelCellFlags файла (проверено по CSheetDWord/SFM_*), поэтому
// маска переносится 1:1: 0x80000000=Text, 0x40000000=Value, 0x200000=Data,
// 0x400000=TextOrientation; wFontSize — знаковое (-4*пт) — без пересчёта.
//
// ЗАКРЫТАЯ МОДЕЛЬ: у MoxelRow.Parent и Section.{Begin,End,Level,Name} нет
// публичного доступа — выставляются рефлексией (мягко, с диагностикой в
// Report). Если сделать их публичными или добавить конструктор — рефлексию
// можно убрать (поиск по "TrySetPrivateField").
//
// НЕ извлекается (v1): бинарные payload картинок (m_pPicture → CPictureHolder7,
// раскладки нет в MOXEL.H) и OLE-хранилищ (m_pContainerItem → CSheetCntrItem :
// COleClientItem+0x78; путь через IOleObject→IPersistStorage→OleSave — v2).
// Геометрия/тип/Z-order рисунков переносятся полностью.
//
// Использование:
//   var table = new V7Table();
//   table.SetTable(tbl);                       // COM-объект Таблица 1С
//   string report;
//   Moxel moxel = MoxelMemLoader.LoadFromMemory(table, out report);
//   moxel.SaveAs(@"C:\out.xlsx", SaveFormat.Excel);
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using static Moxel.Moxel;

namespace Moxel
{
    public sealed class MoxelMemLoader
    {
        // ====================================================================
        //  Смещения (дублируют MemSheet.cs — держать синхронно; по MOXEL.H)
        // ====================================================================

        // --- CSortArray (0x2C): m_indexes CArray @+0x04, m_values CArray @+0x18
        const int SORT_IDX_PDATA = 0x08;
        const int SORT_IDX_NSIZE = 0x0C;
        const int SORT_VAL_PDATA = 0x1C;
        const int SORT_VAL_NSIZE = 0x20;

        // --- CSheet (0x220)
        const int SH_COLUMNS    = 0x2C;   // CSheetFormatsArray CSortArray<int, CSheetFormat*>
        const int SH_ROWS       = 0x58;   // CSheetRowsArray    CSortArray<int, CSheetRow*>
        const int SH_FONTS      = 0x84;   // CSheetFontsArray   CSortArray<int, LOGFONT по значению>
        const int SH_MASKS      = 0xB0;   // CSheetMasksArray   (в модель Moxel не переносится)
        const int SH_HEADER     = 0xDC;   // CSheetCell (встроен)
        const int SH_FOOTER     = 0x120;
        const int SH_NCOLS      = 0x164;
        const int SH_NROWS      = 0x168;
        const int SH_NAMES      = 0x16C;  // CSheetNames (SGI-хэш, 0x18)
        const int SH_PB1_DATA   = 0x188;  // CArray<int> #1 → горизонтальные разрывы (гипотеза)
        const int SH_PB1_SIZE   = 0x18C;
        const int SH_PB2_DATA   = 0x19C;  // CArray<int> #2 → вертикальные разрывы (гипотеза)
        const int SH_PB2_SIZE   = 0x1A0;
        const int SH_DWG_HEAD   = 0x1B0;  // CList<CSheetDrawing*>::m_pNodeHead
        const int SH_DWG_COUNT  = 0x1B8;  // ::m_nCount
        const int SH_SEC_HORZ   = 0x1CC;  // CArray<CSheetOutline>
        const int SH_SEC_VERT   = 0x1E0;
        const int SH_UNKARR2    = 0x1F4;  // CPtrArray → CRect* (кандидат объединений)
        const int SH_UNKARR3    = 0x208;  // CPtrArray → CRect* (кандидат объединений)

        // --- CSheetRow (0x50): base CSheetFormat 0x24, m_obj1 @0x24, m_Keys @0x28, m_Cells @0x3C
        const int ROW_KEYS_PDATA  = 0x2C;
        const int ROW_KEYS_NSIZE  = 0x30;
        const int ROW_CELLS_PDATA = 0x40;
        const int ROW_CELLS_NSIZE = 0x44;
        const int SIZE_ROW = 0x50;

        // --- CSheetFormat (0x24) — байты файла = память-4 (vftable)
        const int FMT_MASK       = 0x04;
        const int FMT_H_W        = 0x08;  // WORD m_Height (union файла: wShow/wColumnPosition/wHeight)
        const int FMT_W_SP       = 0x0A;  // WORD m_Width  (union файла: wStartPage/wWidth/wRowPosition)
        const int FMT_FONTNUM    = 0x0C;
        const int FMT_FONTSIZE   = 0x0E;  // знаковое, = -4*размер_пт — как в файле
        const int FMT_BOLD       = 0x10;  // clFontWeight: 0=Empty, 4=Normal, 7=Bold
        const int FMT_ITALIC     = 0x11;
        const int FMT_UNDERLINE  = 0x12;
        const int FMT_HALIGN     = 0x13;
        const int FMT_VALIGN     = 0x14;
        const int FMT_PATTERN    = 0x15;
        const int FMT_BORDERL    = 0x16;  // union: LineStyle
        const int FMT_BORDERT    = 0x17;  // union: LineWeight
        const int FMT_BORDERR    = 0x18;  // union: Borders
        const int FMT_BORDERB    = 0x19;  // union: TypeOut
        const int FMT_PATTCOLOR  = 0x1A;
        const int FMT_BORDERCOLOR = 0x1B;
        const int FMT_TEXTCOLOR  = 0x1C;
        const int FMT_BKCOLOR    = 0x1D;
        const int FMT_WRAP       = 0x1E;  // bControlContent
        const int FMT_DATAFORMAT = 0x1F;  // bType
        const int FMT_UNPROTECT  = 0x20;  // bAllowEdit
        const int FMT_UNK        = 0x21;  // bXZ1
        const int FMT_TEXTDIR    = 0x22;  // WORD — в файле v7 идёт после 30 байт формата
        const int SIZE_FMT = 0x24;

        // --- CSheetCell (0x44): base CSheetFormat + строки + CByteArray
        const int CELL_TEXT     = 0x24;
        const int CELL_DETAILS  = 0x28;   // «Значение»/расшифровка
        const int CELL_STR3     = 0x2C;   // в файл не сериализуется
        const int CELL_ARR_PDATA = 0x34;
        const int CELL_ARR_NSIZE = 0x38;
        const int SIZE_CELL = 0x44;

        // --- CSheetDrawing (0x74) : CSheetCell + хвост
        const int DWG_PCONTAINER = 0x44;
        const int DWG_PPICTURE   = 0x48;
        const int DWG_TYPE       = 0x4C;
        const int DWG_TL_X       = 0x50;  // SHEETPOINT {xCell, yCell, xOff, yOff}
        const int DWG_TL_Y       = 0x54;
        const int DWG_TL_XO      = 0x58;
        const int DWG_TL_YO      = 0x5C;
        const int DWG_BR_X       = 0x60;
        const int DWG_BR_Y       = 0x64;
        const int DWG_BR_XO      = 0x68;
        const int DWG_BR_YO      = 0x6C;
        const int DWG_INDEX      = 0x70;
        const int SIZE_DWG = 0x74;

        // --- LOGFONTA (0x3C) — поля читает Marshal.PtrToStructure(Moxel.LOGFONT)
        const int SIZE_LOGFONTA = 0x3C;

        // --- CSheetOutline (0x14): {vft, Start, End, m_data, CString Name}
        const int OUT_START = 0x04;
        const int OUT_END   = 0x08;
        const int OUT_LEVEL = 0x0C;
        const int OUT_NAME  = 0x10;
        const int SIZE_OUTLINE = 0x14;

        // --- CSheetNames @0x16C (0x18): vft + SGI __hashtable (VC6)
        const int NAM_BKT_START   = 0x04;   // vector._M_start
        const int NAM_BKT_FINISH  = 0x08;   // vector._M_finish
        const int NAM_NUM_ELEMENTS = 0x10;
        // __hashtable_node { _M_next@0; pair{ CString key@4; CSheetNamedItem val@8 } }
        const int NODE_NEXT = 0x00;
        const int NODE_KEY  = 0x04;
        const int NODE_VAL  = 0x08;
        // --- CSheetNamedItem (0x24, от начала value):
        const int NITEM_TYPE   = 0x00;   // sntCell=1 / sntDraw=2
        const int NITEM_DRAWID = 0x04;
        const int NITEM_SELTYPE = 0x0C;  // внутри вложенного CSheetSelection (после его vft)
        const int NITEM_C1     = 0x10;
        const int NITEM_R1     = 0x14;
        const int NITEM_C2     = 0x18;
        const int NITEM_R2     = 0x1C;
        const int SIZE_NITEM = 0x24;
        const int SIZE_NODE = NODE_VAL + SIZE_NITEM;   // 0x2C

        // --- биты маски (совпадают с MemSheet.SFM_* и MoxelCellFlags)
        const uint MSK_TEXT     = 0x80000000u;
        const uint MSK_VALUE    = 0x40000000u;
        const uint MSK_DATA     = 0x00200000u;
        const uint MSK_TEXTDIR  = 0x00400000u;

        // --- предохранители
        const int MAX_FONTS       = 4096;
        const int MAX_KEYS        = 1 << 20;
        const int MAX_DRAWINGS    = 100000;
        const int MAX_SECTIONS    = 65536;
        const int MAX_UNIONS      = 100000;
        const int MAX_BREAKS      = 100000;
        const int MAX_BUCKETS     = 1 << 16;
        const int MAX_NAMES       = 100000;
        const int MAX_NODE_CHAIN  = 256;
        const int MAX_DATA_BYTES  = 16 * 1024 * 1024;

        // ====================================================================
        //  Состояние и публичный API
        // ====================================================================

        readonly MemSheet _mem;
        readonly IntPtr _pSheet;
        readonly Moxel _mx;
        readonly StringBuilder _log = new StringBuilder();

        int _warnTextFlag, _warnValueFlag, _warnDataFlag;
        int _skipRows, _skipCells, _skipUnions, _skipFonts;

        MoxelMemLoader(MemSheet mem, Moxel target)
        {
            _mem = mem;
            _pSheet = mem.SheetPointer;
            _mx = target;
        }

        /// <summary>Основной путь: по V7Table (использует table.Mem).</summary>
        public static Moxel LoadFromMemory(V7Table table, out string report)
        {
            if (table == null) throw new ArgumentNullException("table");
            return LoadFromMemory(table.Mem, out report);
        }

        public static Moxel LoadFromMemory(MemSheet mem, out string report)
        {
            if (mem == null) throw new ArgumentNullException("mem");
            var mx = new Moxel();
            var loader = new MoxelMemLoader(mem, mx);
            loader.Load();
            report = loader._log.ToString();
            return mx;
        }

        public static Moxel LoadFromMemory(V7Table table)
        {
            string report;
            return LoadFromMemory(table, out report);
        }

        public static Moxel LoadFromMemory(MemSheet mem)
        {
            string report;
            return LoadFromMemory(mem, out report);
        }

        /// <summary>По голому CSheetDoc* (держатель RCW — на вызывающем).</summary>
        public static Moxel LoadFromMemory(IntPtr pSheetDoc, object comObjectHolder)
        {
            string report;
            return LoadFromMemory(new MemSheet(pSheetDoc, comObjectHolder), out report);
        }

        // ====================================================================
        //  Конвейер загрузки
        // ====================================================================

        void Load()
        {
            _log.AppendLine(string.Format(
                "=== MoxelMemLoader: CSheetDoc=0x{0:X8}, CSheet=0x{1:X8}, valid={2} ===",
                _mem.DocPointer.ToInt32(), _pSheet.ToInt32(), _mem.IsValid));

            if (!_mem.IsValid)
                throw new InvalidOperationException(
                    "MoxelMemLoader: CSheet недоступен или back-ptr не сходится (сначала V7Table.SetTable).");

            // Память всегда несёт m_TextDirection — это раскладка v7 (30+2 байта формата)
            _mx.Version = 7;

            _mx.nAllColumnCount = ClampCount(_mem.ColsCount);
            _mx.nAllRowCount = ClampCount(_mem.RowsCount);

            // Формат по умолчанию — базовый CSheetFormat листа
            _mx.DefFormat = new DataCell(ReadFormat(_pSheet), _mx);

            LoadFonts();
            LoadColumns();
            LoadRows();
            LoadHeaderFooter();
            LoadDrawings();
            LoadUnions();
            LoadSections();
            LoadPageBreaks();
            LoadAreaNames();

            _mx.nAllObjectsCount = _mx.Objects.Count;

            _log.AppendLine(string.Format(
                "Итог: cols={0} rows={1} fonts={2} colFormats={3} rowObjs={4} objects={5} unions={6} secH={7} secV={8} breakH={9} breakV={10} names={11}",
                _mx.nAllColumnCount, _mx.nAllRowCount, _mx.FontList.Count, _mx.Columns.Count,
                _mx.Rows.Count, _mx.Objects.Count, _mx.Unions.Count,
                _mx.HorisontalSections.Count, _mx.VerticalSections.Count,
                _mx.HorisontalPageBreaks.Length, _mx.VerticalPageBreaks.Length,
                _mx.AreaNames.Count));

            if (_warnTextFlag + _warnValueFlag + _warnDataFlag > 0)
                _log.AppendLine(string.Format(
                    "Расхождение флагов (бит отсутствовал, но данные есть — бит выставлен): Text={0}, Value={1}, Data={2}",
                    _warnTextFlag, _warnValueFlag, _warnDataFlag));
            if (_skipRows + _skipCells + _skipUnions + _skipFonts > 0)
                _log.AppendLine(string.Format(
                    "Пропущено: строк={0}, ячеек={1}, объединений={2}, шрифтов={3}",
                    _skipRows, _skipCells, _skipUnions, _skipFonts));
        }

        // ====================================================================
        //  Формат и ячейка (файловая семантика DataCell)
        // ====================================================================

        /// <summary>
        /// CSheetFormat* → файловый CSheetFormat (30 байт): сырая маска = dwFlags,
        /// байты копируются все — как пишет родной сериализатор. wShow/wHeight/
        /// wColumnPosition и wStartPage/wWidth/wRowPosition — union'ы файла,
        /// в памяти лежат в тех же байтах, поэтому переносятся парой присваиваний
        /// wHeight/wWidth без дополнительной логики.
        /// </summary>
        CSheetFormat ReadFormat(IntPtr pFmt)
        {
            CSheetFormat f = CSheetFormat.Empty;
            if (pFmt == IntPtr.Zero || !MemSheet.IsReadable(pFmt, SIZE_FMT))
                return f;

            f.dwFlags = unchecked((MoxelCellFlags)(uint)ReadI32(pFmt, FMT_MASK));
            f.wHeight = (short)ReadU16(pFmt, FMT_H_W);
            f.wWidth = (short)ReadU16(pFmt, FMT_W_SP);
            f.wFontNumber = (short)ReadU16(pFmt, FMT_FONTNUM);
            f.wFontSize = (short)ReadU16(pFmt, FMT_FONTSIZE);
            f.bFontBold = (clFontWeight)ReadU8(pFmt, FMT_BOLD);
            f.bFontItalic = ReadU8(pFmt, FMT_ITALIC) != 0;
            f.bFontUnderline = ReadU8(pFmt, FMT_UNDERLINE) != 0;
            f.bHorAlign = (TextHorzAlign)ReadU8(pFmt, FMT_HALIGN);
            f.bVertAlign = (TextVertAlign)ReadU8(pFmt, FMT_VALIGN);
            f.bPatternType = ReadU8(pFmt, FMT_PATTERN);
            f.bBorderLeft = (BorderStyle)ReadU8(pFmt, FMT_BORDERL);
            f.bBorderTop = (BorderStyle)ReadU8(pFmt, FMT_BORDERT);
            f.bBorderRight = (BorderStyle)ReadU8(pFmt, FMT_BORDERR);
            f.bBorderBottom = (BorderStyle)ReadU8(pFmt, FMT_BORDERB);
            f.bPatternColor = ReadU8(pFmt, FMT_PATTCOLOR);
            f.bBorderColor = ReadU8(pFmt, FMT_BORDERCOLOR);
            f.bFontColor = ReadU8(pFmt, FMT_TEXTCOLOR);
            f.bBackground = ReadU8(pFmt, FMT_BKCOLOR);
            f.bControlContent = (TextControl)ReadU8(pFmt, FMT_WRAP);
            f.bType = (ContentType)ReadU8(pFmt, FMT_DATAFORMAT);
            f.bAllowEdit = ReadU8(pFmt, FMT_UNPROTECT) != 0;
            f.bXZ1 = ReadU8(pFmt, FMT_UNK);
            return f;
        }

        /// <summary>
        /// CSheetCell* → DataCell (файловая семантика): Text/Value/Data читаются
        /// по битам маски; страховка — если строка в памяти не пуста, а бит не
        /// выставлен, данные переносятся и бит добавляется (счётчик в Report).
        /// </summary>
        DataCell BuildDataCell(IntPtr pCell)
        {
            CSheetFormat fmt = ReadFormat(pCell);
            uint mask = unchecked((uint)fmt.dwFlags);

            string text = ReadCString(ReadPtr(pCell, CELL_TEXT));
            if ((mask & MSK_TEXT) != 0 || text.Length > 0)
            {
                if ((mask & MSK_TEXT) == 0) { _warnTextFlag++; fmt.dwFlags |= MoxelCellFlags.Text; }
            }

            string val = ReadCString(ReadPtr(pCell, CELL_DETAILS));
            if ((mask & MSK_VALUE) != 0 || val.Length > 0)
            {
                if ((mask & MSK_VALUE) == 0) { _warnValueFlag++; fmt.dwFlags |= MoxelCellFlags.Value; }
            }

            int nData = ReadI32(pCell, CELL_ARR_NSIZE);
            byte[] data = null;
            if ((mask & MSK_DATA) != 0 || nData > 0)
            {
                data = ReadBytes(pCell);
                if (nData > 0 && (mask & MSK_DATA) == 0) { _warnDataFlag++; fmt.dwFlags |= MoxelCellFlags.Data; }
            }

            if ((mask & MSK_TEXTDIR) != 0)
                fmt.dwFlags |= MoxelCellFlags.TextOrientation;

            var dc = new DataCell(fmt, _mx);
            if ((fmt.dwFlags & MoxelCellFlags.Text) != 0) dc.Text = text;
            if ((fmt.dwFlags & MoxelCellFlags.Value) != 0) dc.Value = val;
            if ((fmt.dwFlags & MoxelCellFlags.Data) != 0) dc.Data = data ?? new byte[0];
            if ((fmt.dwFlags & MoxelCellFlags.TextOrientation) != 0)
                dc.TextOrientation = (short)ReadU16(pCell, FMT_TEXTDIR);
            return dc;
        }

        // ====================================================================
        //  Шрифты / колонки / строки / колонтитулы
        // ====================================================================

        void LoadFonts()
        {
            IntPtr keys, vals; int nKeys, nVals;
            if (!GetSortArray(SH_FONTS, out keys, out nKeys, out vals, out nVals, "fonts"))
            {
                _log.AppendLine("Шрифты: пусто");
                return;
            }

            int n = Math.Min(nKeys, nVals);
            if (n > MAX_FONTS)
            {
                _log.AppendLine(string.Format("Шрифты: {0} > {1}, усечено", n, MAX_FONTS));
                n = MAX_FONTS;
            }
            if (!CheckRange(keys, n * 4, "fonts.keys") || !CheckRange(vals, n * SIZE_LOGFONTA, "fonts.vals"))
                return;

            for (int i = 0; i < n; i++)
            {
                IntPtr lf = vals + i * SIZE_LOGFONTA;
                if (!MemSheet.IsReadable(lf, SIZE_LOGFONTA)) { _skipFonts++; continue; }
                int key = ReadI32(keys, i * 4);   // ключ CSortArray = номер шрифта в кэше
                try
                {
                    var font = (Moxel.LOGFONT)Marshal.PtrToStructure(lf, typeof(Moxel.LOGFONT));
                    _mx.FontList[key] = font;
                }
                catch (Exception ex)
                {
                    _skipFonts++;
                    _log.AppendLine(string.Format("  ! font[{0}]: PtrToStructure: {1}", i, ex.Message));
                }
            }
            _log.AppendLine(string.Format("Шрифты: {0} из {1}", _mx.FontList.Count, n));
        }

        void LoadColumns()
        {
            IntPtr keys, vals; int nKeys, nVals;
            if (!GetSortArray(SH_COLUMNS, out keys, out nKeys, out vals, out nVals, "columns"))
            {
                _log.AppendLine("Форматы колонок: пусто");
                return;
            }

            int n = Math.Min(nKeys, nVals);
            if (!CheckRange(keys, n * 4, "columns.keys") || !CheckRange(vals, n * 4, "columns.vals"))
                return;

            for (int i = 0; i < n; i++)
            {
                int col = ReadI32(keys, i * 4);
                IntPtr pFmt = ReadPtr(vals, i * 4);
                if (pFmt == IntPtr.Zero || !MemSheet.IsReadable(pFmt, SIZE_FMT)) { _skipCells++; continue; }
                _mx.Columns[col] = new DataCell(ReadFormat(pFmt), _mx);
            }
            _log.AppendLine(string.Format("Форматы колонок: {0} из {1}", _mx.Columns.Count, n));
        }

        void LoadRows()
        {
            IntPtr keys, vals; int nKeys, nVals;
            if (!GetSortArray(SH_ROWS, out keys, out nKeys, out vals, out nVals, "rows"))
            {
                _log.AppendLine("Строки: пусто");
                return;
            }

            int n = Math.Min(Math.Min(nKeys, nVals), MAX_KEYS);
            if (!CheckRange(keys, n * 4, "rows.keys") || !CheckRange(vals, n * 4, "rows.vals"))
                return;

            for (int i = 0; i < n; i++)
            {
                int rowKey = ReadI32(keys, i * 4);
                IntPtr pRow = ReadPtr(vals, i * 4);
                if (pRow == IntPtr.Zero || !MemSheet.IsReadable(pRow, SIZE_ROW)) { _skipRows++; continue; }

                var row = new MoxelRow();
                row.FormatCell = ReadFormat(pRow);
                row.values = new Dictionary<int, DataCell>();
                TrySetPrivateField(row, "Parent", _mx, "MoxelRow.Parent");

                IntPtr ck = ReadPtr(pRow, ROW_KEYS_PDATA);
                int ckn = ReadI32(pRow, ROW_KEYS_NSIZE);
                IntPtr cv = ReadPtr(pRow, ROW_CELLS_PDATA);
                int cvn = ReadI32(pRow, ROW_CELLS_NSIZE);
                int m = Math.Min(ckn, cvn);
                if (m > 0 && ck != IntPtr.Zero && cv != IntPtr.Zero &&
                    CheckRange(ck, m * 4, "row.cells.keys") && CheckRange(cv, m * 4, "row.cells.vals"))
                {
                    for (int j = 0; j < m; j++)
                    {
                        int colKey = ReadI32(ck, j * 4);
                        IntPtr pCell = ReadPtr(cv, j * 4);
                        if (pCell == IntPtr.Zero || !MemSheet.IsReadable(pCell, SIZE_CELL)) { _skipCells++; continue; }
                        row.values[colKey] = BuildDataCell(pCell);
                    }
                }
                _mx.Rows[rowKey] = row;
            }
            _log.AppendLine(string.Format("Строки: {0} объектов из {1} (ячеек пропущено: {2})",
                _mx.Rows.Count, n, _skipCells));
        }

        void LoadHeaderFooter()
        {
            _mx.Header = BuildDataCell(_pSheet + SH_HEADER);
            _mx.Footer = BuildDataCell(_pSheet + SH_FOOTER);
            _log.AppendLine(string.Format("Колонтитулы: верх='{0}', низ='{1}' (wShow в wHeight: 1=показ, 0xFFFF=скрыт)",
                Clip(_mx.Header.Text, 60), Clip(_mx.Footer.Text, 60)));
        }

        // ====================================================================
        //  Рисунки (CList<CSheetDrawing*>)
        // ====================================================================

        void LoadDrawings()
        {
            IntPtr node = ReadPtr(_pSheet, SH_DWG_HEAD);
            int count = ReadI32(_pSheet, SH_DWG_COUNT);
            if (node == IntPtr.Zero || count <= 0)
            {
                _log.AppendLine("Рисунки: нет");
                return;
            }
            if (count > MAX_DRAWINGS)
            {
                _log.AppendLine(string.Format("Рисунки: {0} > {1}, усечено", count, MAX_DRAWINGS));
                count = MAX_DRAWINGS;
            }

            for (int i = 0; i < count && node != IntPtr.Zero; i++)
            {
                IntPtr pDwg = ReadPtr(node, 8);   // CList::CNode {pNext@0, pPrev@4, data@8}
                if (pDwg != IntPtr.Zero && MemSheet.IsReadable(pDwg, SIZE_DWG))
                {
                    var eo = new EmbeddedObject();
                    eo.Parent = _mx;

                    CSheetFormat fmt = ReadFormat(pDwg);
                    uint mask = unchecked((uint)fmt.dwFlags);
                    string text = ReadCString(ReadPtr(pDwg, CELL_TEXT));
                    string val = ReadCString(ReadPtr(pDwg, CELL_DETAILS));
                    if ((mask & MSK_TEXT) != 0 || text.Length > 0)
                    {
                        if ((mask & MSK_TEXT) == 0) { _warnTextFlag++; fmt.dwFlags |= MoxelCellFlags.Text; }
                    }
                    if ((mask & MSK_VALUE) != 0 || val.Length > 0)
                    {
                        if ((mask & MSK_VALUE) == 0) { _warnValueFlag++; fmt.dwFlags |= MoxelCellFlags.Value; }
                    }
                    if ((mask & MSK_TEXTDIR) != 0)
                        fmt.dwFlags |= MoxelCellFlags.TextOrientation;

                    eo.FormatCell = fmt;
                    if ((fmt.dwFlags & MoxelCellFlags.Text) != 0) eo.Text = text;
                    if ((fmt.dwFlags & MoxelCellFlags.Value) != 0) eo.Value = val;
                    if ((fmt.dwFlags & MoxelCellFlags.Data) != 0) eo.Data = ReadBytes(pDwg);
                    if ((fmt.dwFlags & MoxelCellFlags.TextOrientation) != 0)
                        eo.TextOrientation = (short)ReadU16(pDwg, FMT_TEXTDIR);

                    var pic = new Picture();
                    pic.dwType = (ObjectType)ReadI32(pDwg, DWG_TYPE);
                    pic.dwColumnStart = ReadI32(pDwg, DWG_TL_X);
                    pic.dwRowStart = ReadI32(pDwg, DWG_TL_Y);
                    pic.dwOffsetLeft = ReadI32(pDwg, DWG_TL_XO);
                    pic.dwOffsetTop = ReadI32(pDwg, DWG_TL_YO);
                    pic.dwColumnEnd = ReadI32(pDwg, DWG_BR_X);
                    pic.dwRowEnd = ReadI32(pDwg, DWG_BR_Y);
                    pic.dwOffsetRight = ReadI32(pDwg, DWG_BR_XO);
                    pic.dwOffsetBottom = ReadI32(pDwg, DWG_BR_YO);
                    pic.dwZOrder = ReadI32(pDwg, DWG_INDEX);
                    eo.Picture = pic;
                    // eo.pObject / eo.OleObjectStorage: payload картинок и OLE
                    // из памяти не извлекается (v1) — см. шапку файла

                    _mx.Objects.Add(eo);
                }
                node = ReadPtr(node, 0);          // pNext
            }
            _log.AppendLine(string.Format("Рисунки: {0}", _mx.Objects.Count));
        }

        // ====================================================================
        //  Объединения (UnkArr2/UnkArr3 — CPtrArray → CRect*, гипотеза)
        // ====================================================================

        void LoadUnions()
        {
            AddUnionsFrom(SH_UNKARR2, "UnkArr2");
            AddUnionsFrom(SH_UNKARR3, "UnkArr3");
            _log.AppendLine(string.Format("Объединения ячеек: {0} (отброшено по валидации: {1})",
                _mx.Unions.Count, _skipUnions));
        }

        void AddUnionsFrom(int baseOff, string what)
        {
            IntPtr pData = ReadPtr(_pSheet, baseOff + 4);   // CPtrArray: vft@0, pData@+4, nSize@+8
            int n = ReadI32(_pSheet, baseOff + 8);
            if (pData == IntPtr.Zero || n <= 0)
            {
                _log.AppendLine(string.Format("{0}: пусто", what));
                return;
            }
            if (n > MAX_UNIONS)
            {
                _log.AppendLine(string.Format("{0}: {1} > {2}, усечено", what, n, MAX_UNIONS));
                n = MAX_UNIONS;
            }
            if (!CheckRange(pData, n * 4, what)) return;

            bool checkBounds = _mx.nAllColumnCount > 0 && _mx.nAllRowCount > 0;

            for (int i = 0; i < n; i++)
            {
                IntPtr pRect = ReadPtr(pData, i * 4);
                if (pRect == IntPtr.Zero || !MemSheet.IsReadable(pRect, 16)) { _skipUnions++; continue; }

                int l = ReadI32(pRect, 0);
                int t = ReadI32(pRect, 4);
                int r = ReadI32(pRect, 8);
                int b = ReadI32(pRect, 12);

                if (l > r) { int tmp = l; l = r; r = tmp; }
                if (t > b) { int tmp = t; t = b; b = tmp; }

                // настоящее объединение минимум 2 ячейки и внутри таблицы
                bool realMerge = (r > l) || (b > t);
                bool inside = l >= 0 && t >= 0 &&
                    r < _mx.nAllColumnCount && b < _mx.nAllRowCount;
                if (!realMerge || (checkBounds && !inside)) { _skipUnions++; continue; }

                var u = new CellsUnion();
                u.dwLeft = l; u.dwTop = t; u.dwRight = r; u.dwBottom = b;
                _mx.Unions.Add(u);
            }
        }

        // ====================================================================
        //  Секции и разрывы страниц
        // ====================================================================

        void LoadSections()
        {
            _mx.HorisontalSections = ReadSectionArray(SH_SEC_HORZ, "горизонтальные");
            _mx.VerticalSections = ReadSectionArray(SH_SEC_VERT, "вертикальные");
        }

        List<Section> ReadSectionArray(int baseOff, string what)
        {
            var res = new List<Section>();
            IntPtr pData = ReadPtr(_pSheet, baseOff + 4);   // CArray: vft@0, pData@+4, nSize@+8
            int n = ReadI32(_pSheet, baseOff + 8);
            if (pData == IntPtr.Zero || n <= 0)
            {
                _log.AppendLine(string.Format("Секции {0}: пусто", what));
                return res;
            }
            if (n > MAX_SECTIONS)
            {
                _log.AppendLine(string.Format("Секции {0}: {1} > {2}, усечено", what, n, MAX_SECTIONS));
                n = MAX_SECTIONS;
            }
            if (!CheckRange(pData, n * SIZE_OUTLINE, what)) return res;

            for (int i = 0; i < n; i++)
            {
                IntPtr e = pData + i * SIZE_OUTLINE;
                var s = new Section();
                // m_data@0xC → Level: гипотеза по позиции в файловой Section{Begin,End,Level,Name}
                TrySetPrivateField(s, "Begin", ReadI32(e, OUT_START), "Section.Begin");
                TrySetPrivateField(s, "End", ReadI32(e, OUT_END), "Section.End");
                TrySetPrivateField(s, "Level", ReadI32(e, OUT_LEVEL), "Section.Level");
                TrySetPrivateField(s, "Name", ReadCString(ReadPtr(e, OUT_NAME)), "Section.Name");
                res.Add(s);
            }
            _log.AppendLine(string.Format("Секции {0}: {1}", what, res.Count));
            return res;
        }

        void LoadPageBreaks()
        {
            _mx.HorisontalPageBreaks = ReadIntCArray(SH_PB1_DATA, SH_PB1_SIZE, "разрывы H") ?? new int[0];
            _mx.VerticalPageBreaks = ReadIntCArray(SH_PB2_DATA, SH_PB2_SIZE, "разрывы V") ?? new int[0];
            _log.AppendLine(string.Format(
                "Разрывы страниц: H={0}, V={1} (гипотеза PB1=H/PB2=V — сверить на живой таблице)",
                _mx.HorisontalPageBreaks.Length, _mx.VerticalPageBreaks.Length));
        }

        int[] ReadIntCArray(int offData, int offSize, string what)
        {
            IntPtr pData = ReadPtr(_pSheet, offData);
            int n = ReadI32(_pSheet, offSize);
            if (pData == IntPtr.Zero || n <= 0) return null;
            if (n > MAX_BREAKS)
            {
                _log.AppendLine(string.Format("{0}: {1} > {2}, усечено", what, n, MAX_BREAKS));
                n = MAX_BREAKS;
            }
            if (!CheckRange(pData, n * 4, what)) return null;

            var res = new int[n];
            for (int i = 0; i < n; i++)
                res[i] = ReadI32(pData, i * 4);
            return res;
        }

        // ====================================================================
        //  Именованные области (CSheetNames = SGI-хэш VC6)
        // ====================================================================

        void LoadAreaNames()
        {
            IntPtr names = _pSheet + SH_NAMES;   // CSheetNames (0x18), встроен в CSheet
            IntPtr bktStart = ReadPtr(names, NAM_BKT_START);
            IntPtr bktFinish = ReadPtr(names, NAM_BKT_FINISH);
            int numElements = ReadI32(names, NAM_NUM_ELEMENTS);

            if (bktStart == IntPtr.Zero || bktFinish == IntPtr.Zero ||
                bktFinish.ToInt32() < bktStart.ToInt32())
            {
                _log.AppendLine(string.Format(
                    "Именованные области: хэш не опознан (start=0x{0:X8}, finish=0x{1:X8}, n={2})",
                    bktStart.ToInt32(), bktFinish.ToInt32(), numElements));
                return;
            }

            int buckets = (bktFinish.ToInt32() - bktStart.ToInt32()) / 4;
            if (buckets <= 0 || buckets > MAX_BUCKETS || numElements <= 0 || numElements > MAX_NAMES)
            {
                _log.AppendLine(string.Format(
                    "Именованные области: параметры вне допуска (buckets={0}, n={1}) — пропуск",
                    buckets, numElements));
                return;
            }
            if (!CheckRange(bktStart, buckets * 4, "names.buckets")) return;

            for (int b = 0; b < buckets && _mx.AreaNames.Count < numElements; b++)
            {
                IntPtr node = ReadPtr(bktStart, b * 4);
                int chain = 0;
                while (node != IntPtr.Zero && chain < MAX_NODE_CHAIN && _mx.AreaNames.Count < numElements)
                {
                    if (!MemSheet.IsReadable(node, SIZE_NODE)) break;
                    IntPtr next = ReadPtr(node, NODE_NEXT);

                    var area = new MoxelArea();
                    area.Name = ReadCString(ReadPtr(node, NODE_KEY));
                    IntPtr v = node + NODE_VAL;                  // CSheetNamedItem
                    area.Area.Unknown1 = 1;                      // в файлах всегда 1
                    area.Area.Unknown2 = ReadI32(v, NITEM_DRAWID);
                    area.Area.AreaType = ReadI32(v, NITEM_TYPE); // sntCell=1 / sntDraw=2
                    area.Area.ColumnBegin = ReadI32(v, NITEM_C1);
                    area.Area.RowBegin = ReadI32(v, NITEM_R1);
                    area.Area.ColumnEnd = ReadI32(v, NITEM_C2);
                    area.Area.RowEnd = ReadI32(v, NITEM_R2);
                    _mx.AreaNames.Add(area);

                    node = next;
                    chain++;
                }
            }
            _log.AppendLine(string.Format("Именованные области: {0} (buckets={1}, ожидалось {2})",
                _mx.AreaNames.Count, buckets, numElements));
        }

        // ====================================================================
        //  Вспомогательные
        // ====================================================================

        /// <summary>Прочитать параллельные массивы CSortArray (keys+values).</summary>
        bool GetSortArray(int baseOff, out IntPtr keys, out int nKeys,
                          out IntPtr vals, out int nVals, string what)
        {
            keys = ReadPtr(_pSheet, baseOff + SORT_IDX_PDATA);
            nKeys = ReadI32(_pSheet, baseOff + SORT_IDX_NSIZE);
            vals = ReadPtr(_pSheet, baseOff + SORT_VAL_PDATA);
            nVals = ReadI32(_pSheet, baseOff + SORT_VAL_NSIZE);
            if (nKeys <= 0 || nVals <= 0 || keys == IntPtr.Zero || vals == IntPtr.Zero)
                return false;
            return true;
        }

        bool CheckRange(IntPtr p, int byteSize, string what)
        {
            if (byteSize <= 0) return true;
            if (!MemSheet.IsReadable(p, byteSize))
            {
                _log.AppendLine(string.Format("  ! {0}: память 0x{1:X8}+{2}B недоступна — пропуск",
                    what, p.ToInt32(), byteSize));
                return false;
            }
            return true;
        }

        /// <summary>
        /// Мягкая установка закрытого поля модели (MoxelRow.Parent,
        /// Section.{Begin,End,Level,Name}). Не падает — пишет в Report.
        /// </summary>
        static void TrySetPrivateField(object obj, string fieldName, object value, string what)
        {
            if (obj == null) return;
            try
            {
                FieldInfo fi = obj.GetType().GetField(fieldName,
                    BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);
                if (fi == null)
                {
                    Debug.WriteLine(string.Format("MoxelMemLoader: {0}: поле '{1}' не найдено", what, fieldName));
                    return;
                }
                fi.SetValue(obj, value);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(string.Format("MoxelMemLoader: {0}: {1}", what, ex.Message));
            }
        }

        static int ClampCount(int v)
        {
            if (v < 0) return 0;
            if (v > MAX_KEYS) return MAX_KEYS;
            return v;
        }

        static string Clip(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace("\r\n", "\\n").Replace("\r", "\\n").Replace("\n", "\\n").Replace("\t", "\\t");
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }

        byte[] ReadBytes(IntPtr pCell)
        {
            IntPtr pData = ReadPtr(pCell, CELL_ARR_PDATA);
            int n = ReadI32(pCell, CELL_ARR_NSIZE);
            if (pData == IntPtr.Zero || n <= 0) return new byte[0];
            if (n > MAX_DATA_BYTES)
            {
                _log.AppendLine(string.Format("  ! данные ячейки {0} байт — усечено до {1}", n, MAX_DATA_BYTES));
                n = MAX_DATA_BYTES;
            }
            if (!CheckRange(pData, n, "cell.data")) return new byte[0];

            var buf = new byte[n];
            Marshal.Copy(pData, buf, 0, n);
            return buf;
        }

        string ReadCString(IntPtr pch)
        {
            return MemSheet.ReadCString(pch);
        }

        // --- безопасные чтения (копии MemSheet, те же защиты IsBadReadPtr) ---

        static IntPtr ReadPtr(IntPtr p, int off)
        {
            if (!MemSheet.IsReadable(p + off, 4)) return IntPtr.Zero;
            return Marshal.ReadIntPtr(p, off);
        }

        static int ReadI32(IntPtr p, int off)
        {
            if (!MemSheet.IsReadable(p + off, 4)) return 0;
            return Marshal.ReadInt32(p, off);
        }

        static ushort ReadU16(IntPtr p, int off)
        {
            if (!MemSheet.IsReadable(p + off, 2)) return 0;
            return (ushort)Marshal.ReadInt16(p, off);
        }

        static byte ReadU8(IntPtr p, int off)
        {
            if (!MemSheet.IsReadable(p + off, 1)) return 0;
            return Marshal.ReadByte(p, off);
        }
    }
}
