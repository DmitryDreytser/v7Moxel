// ============================================================================
// MemSheet.cs — ПРЯМОЕ чтение содержимого таблицы 1С 7.7 из памяти процесса.
//
// Альтернатива сериализации (ReadMoxel → CFile::Write → разбор mxl): CSheet
// хранит ячейки в обычных MFC-коллекциях, и их можно вычитать указателями —
// без CArchive, без буфера и без парсинга формата файла.
//
// РАСКЛАДКА (верифицировано по MOXEL.H; 1С 7.7 = ANSI/MFC42/VC6, x86):
//
//   CSheetDoc (MOXEL.H:752)
//     +0x0B0  CSheet m_Sheet            ← pSheet = pDoc + 0xB0
//
//   CSheet : CSheetFormat (MOXEL.H:508, size 0x220)
//     +0x000  vftable                   (CSheetFormat base, см. OFF_FMT_*)
//     +0x004  m_mask (default-формат листа, тот же набор SFM_*)
//     +0x2C   m_ColumnsArray  CSortArray<int, CSheetFormat*>   (0x2C байт)
//     +0x58   m_RowsArray     CSortArray<int, CSheetRow*>
//     +0x84   m_FontsArray    CSortArray<int, LOGFONT(по значению, 0x3C)>
//     +0xB0   m_MasksArray    CSortArray<int, CString(по значению)>
//     +0xDC   m_Header  CSheetCell (колонтитул верхний)
//     +0x120  m_Footer  CSheetCell (колонтитул нижний)
//     +0x164  m_nCols; +0x168 m_nRows
//     +0x1AC  m_SheetDrawingList CList<CSheetDrawing*>
//             → pNodeHead@0x1B0, nCount@0x1B8; CNode{pNext@0,pPrev@4,data@8}
//     +0x1CC  m_HorzSectionArr CArray<CSheetOutline(0x14)>  — секции по строкам
//     +0x1E0  m_VertSectionArr CArray<CSheetOutline>        — секции по колонкам
//     +0x21C  m_pSheetDoc (back-ptr)
//
//   CSortArray (MOXEL.H:6, size 0x2C) — ПАРАЛЛЕЛЬНЫЕ массивы ключ/значение:
//     +0x04 m_indexes CArray<int>      → m_pData@+0x08, m_nSize@+0x0C
//     +0x18 m_values  CArray<T>        → m_pData@+0x1C, m_nSize@+0x20
//
//   CSheetRow : CSheetFormat (MOXEL.H:911, size 0x50) — РАЗРЕЖЕННАЯ строка:
//     +0x004 m_mask (формат строки; m_Height@0x08 — высота)
//     +0x28  m_Keys  CArray<int>         → pData@+0x2C, nSize@+0x30 (индексы колонок)
//     +0x3C  m_Cells CArray<CSheetCell*> → pData@+0x40, nSize@+0x44
//     Отсутствие объекта строки = у строки нет ни текста, ни своего формата.
//
//   CSheetCell : CSheetFormat (MOXEL.H:258, size 0x44):
//     +0x004 m_mask (свойства, выставленные ЯВНО в этой ячейке)
//     +0x24  m_strText    CString (Text, ANSI/CP1251; бит 0x80000000)
//     +0x28  m_strDetails CString (Value; бит 0x40000000)
//     +0x2C  m_strUnk     CString (третья строка — в файл не сериализуется)
//     +0x30  m_arr CByteArray → pData@+0x34, nSize@+0x38
//
//   CSheetFormat (MOXEL.H:153, size 0x24) — база листа/строки/колонки/ячейки:
//     mask@0x04, Height(W)@0x08, Width(W)@0x0A, FontNum(W)@0x0C,
//     FontSize(W)@0x0E (знаковое, = -4*размер_пт), Bold@0x10, Italic@0x11,
//     Underline@0x12, HAlign@0x13, VAlign@0x14, Pattern@0x15,
//     BorderL/LineStyle@0x16, BorderT/LineWeight@0x17, BorderR/Borders@0x18,
//     BorderB/TypeOut@0x19, PatternColor@0x1A, BorderColor@0x1B,
//     TextColor@0x1C, BkColor@0x1D, Wrap@0x1E, DataFormat@0x1F,
//     Unprotected@0x20, TextDirection(W)@0x22
//
//   CString (MFC42/VC6): {LPTSTR m_pchData}; перед данными заголовок
//   CStringData {long nRefs; int nDataLength; int nAllocLength;}
//   → длина = *(int*)(pch-8), символы ANSI по pch.
//
// ОГРАНИЧЕНИЯ / УТОЧНЕНИЯ (сверено с парсером MXL пользователя — Moxel.cs/CSheetFormat.cs):
//   - адресация row/col 0-based (в 1С отображаются 1-based);
//   - биты маски in-memory СОВПАДАЮТ с MoxelCellFlags файла:
//       0x00200000 = Data (байтовые данные), 0x40000000 = Value,
//       0x80000000 = Text; bFontBold: 0=Empty, 4=Normal, 7=Bold (clFontWeight);
//   - colors — индексы палитры (a1CPallete из Enums.cs, совпадает);
//   - файловый CSheetFormat (30 байт, Pack=1) = in-memory CSheetFormat,
//       сдвинутый на -4 (vftable): dwFlags@0x00↔mask@0x04, ...
//       bAllowEdit@0x1C↔Unprotected@0x20, bXZ1@0x1D↔Unknown@0x21,
//       в памяти еще + TextDirection(W)@0x22 (в файле v7 — после формата);
//   - Merge-области: кандидаты UnkArr2/3 (CPtrArray → CRect* — 16 байт
//       left/top/right/bottom = байт-в-байт CellsUnion из парсера);
//   - разрывы страниц: два CArray<int>, точно встающие в зону 0x184..0x1AC.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace Moxel
{
    // ========================================================================
    //  Результат чтения ячейки/формата
    // ========================================================================

    public sealed class SheetCellInfo
    {
        public int Row;
        public int Col;

        /// <summary>Существует ли физический объект CSheetCell (разреженное хранение).</summary>
        public bool HasCell;

        /// <summary>Текст ячейки (m_strText). Для чистых форматов = null.</summary>
        public string Text;
        /// <summary>m_strDetails — «Значение» ячейки (MoxelCellFlags.Value, 0x40000000).</summary>
        public string Str2;
        /// <summary>m_strUnk — третья строка (в файл не сериализуется, runtime).</summary>
        public string Str3;

        /// <summary>CByteArray m_arr.nSize (для чего 1С использует — уточнить дампом).</summary>
        public int BytesCount = -1;

        /// <summary>Суммарная маска эффективного формата (SFM_*).</summary>
        public int Mask;

        // --- поля формата: заполнены ТОЛЬКО если бит маски выставлен ---

        public int? FontIndex;              // индекс в m_FontsArray
        public string FontName;             // resolved из LOGFONTA (если FontIndex задан)
        public short? FontSizeRaw;          // как в памяти
        public float? FontSizePt;           // = FontSizeRaw / -4 (по комментарию MOXEL.H)

        public byte? BoldRaw;  public bool? Bold;
        public byte? ItalicRaw; public bool? Italic;
        public byte? UnderlineRaw; public bool? Underline;

        public ushort? Height;              // высота строки (формат строки)
        public ushort? Width;               // ширина колонки (формат колонки)

        public byte? HAlign;                // 0-лево, 2-право, 4-по ширине, 6-центр
        public byte? VAlign;                // 0-верх, 8-низ, 24-центр (по комментарию MOXEL.H)

        public byte? BorderLeft, BorderTop, BorderRight, BorderBottom;
        public byte? LineStyle, LineWeight, Borders, TypeOut; // union-поля (те же байты)

        public byte? BorderColor, TextColor, BkColor, Pattern, PatternColor;
        public byte? Wrap;                  // 0-авто,1-обрезать,2-забивать,3-переносить,4-красный,5-заб+кр
        public byte? DataFormat;            // 0-текст,1-выражение,2-шаблон,3-фшаблон
        public bool? Unprotected;
        public ushort? TextDirection;
        public bool HasText;                // бит 0x80000000 (MoxelCellFlags.Text)
        public bool HasValue;               // бит 0x40000000 (MoxelCellFlags.Value)
        public bool HasData;                // бит 0x00200000 (MoxelCellFlags.Data)

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(string.Format("R{0}C{1}{2} '{3}'",
                Row, Col, HasCell ? "" : "(объекта нет)", Text ?? ""));
            if (FontName != null) sb.Append(" font='").Append(FontName).Append('\'');
            if (FontSizePt.HasValue) sb.Append(' ').Append(FontSizePt.Value.ToString("0.#"));
            if (Bold == true) sb.Append(" B");
            if (Italic == true) sb.Append(" I");
            if (Underline == true) sb.Append(" U");
            if (HAlign.HasValue) sb.Append(" HA=").Append(HAlign.Value);
            if (VAlign.HasValue) sb.Append(" VA=").Append(VAlign.Value);
            if (DataFormat.HasValue) sb.Append(" DF=").Append(DataFormat.Value);
            if (Wrap.HasValue) sb.Append(" Wrap=").Append(Wrap.Value);
            sb.Append(string.Format(" mask=0x{0:X8}", Mask));
            return sb.ToString();
        }
    }

    public sealed class SheetDrawingInfo
    {
        public int Index;
        public int Type;                    // 1-Line 2-Rect 3-Text 4-OLE 5-Picture
        public int Col1, Row1;              // левая-верхняя ячейка (0-based)
        public int XOff1, YOff1;            // смещение внутри ячейки
        public int Col2, Row2;              // правая-нижняя ячейка
        public int XOff2, YOff2;
        public string Text = "";            // m_strText базового CSheetCell (для Text-боксов)
        public string Str2 = "";
        public IntPtr pPicture;
        public IntPtr pContainer;

        public override string ToString()
        {
            return string.Format("drawing[{0}] type={1} ({2},{3})-({4},{5}) text='{6}'",
                Index, Type, Col1, Row1, Col2, Row2, Text);
        }
    }

    public sealed class SheetSectionInfo
    {
        public bool Vertical;
        public int Start;                   // 0-based
        public int End;
        public string Name = "";
    }

    // ========================================================================
    //  MemSheet — ридер
    // ========================================================================

    public sealed class MemSheet
    {
        // --------------------------- CSheetDoc ------------------------------
        const int OFF_SHEETDOC_SHEET = 0x0B0;

        // ------------------------ CSortArray (0x2C) -------------------------
        const int OFF_SORT_IDX_PDATA = 0x08;
        const int OFF_SORT_IDX_NSIZE = 0x0C;
        const int OFF_SORT_VAL_PDATA = 0x1C;
        const int OFF_SORT_VAL_NSIZE = 0x20;

        // ----------------------- члены CSheet -------------------------------
        const int OFF_SHEET_COLUMNSARRAY   = 0x2C;
        const int OFF_SHEET_ROWSARRAY      = 0x58;
        const int OFF_SHEET_FONTSARRAY     = 0x84;
        const int OFF_SHEET_MASKSARRAY     = 0xB0;
        const int OFF_SHEET_HEADER         = 0xDC;
        const int OFF_SHEET_FOOTER         = 0x120;
        const int OFF_SHEET_NCOLS          = 0x164;
        const int OFF_SHEET_NROWS          = 0x168;
        const int OFF_SHEET_DRAWLIST_HEAD  = 0x1B0;
        const int OFF_SHEET_DRAWLIST_COUNT = 0x1B8;
        const int OFF_SHEET_HORZSECTIONS   = 0x1CC;
        const int OFF_SHEET_VERTSECTIONS   = 0x1E0;
        const int OFF_SHEET_UNKARR2        = 0x1F4;
        const int OFF_SHEET_UNKARR3        = 0x208;

        // Зона 0x184..0x1AC (в RE-дампе помечена как m_Obj5 + m_data2_0..8):
// ДВА CArray<int> встают туда ВСТЫК: [vft,pData,nSize,nMax,grow]@0x184
// и [vft,pData,nSize,nMax,grow]@0x198, конец 0x184+0x28=0x1AC — ровно
// начало списка рисунков. По порядку в Serialize (после секций) — это
// разрывы страниц. Какой H/какой V — сверить на живой таблице.
        const int OFF_SHEET_PB1_PDATA = 0x188;
        const int OFF_SHEET_PB1_NSIZE = 0x18C;
        const int OFF_SHEET_PB2_PDATA = 0x19C;
        const int OFF_SHEET_PB2_NSIZE = 0x1A0;

        // ----------------------- члены CSheetRow ----------------------------
        const int OFF_ROW_KEYS_PDATA  = 0x2C;
        const int OFF_ROW_KEYS_NSIZE  = 0x30;
        const int OFF_ROW_CELLS_PDATA = 0x40;
        const int OFF_ROW_CELLS_NSIZE = 0x44;

        // --------------------- CSheetFormat / CSheetCell --------------------
        const int OFF_FMT_MASK        = 0x04;
        const int OFF_FMT_HEIGHT      = 0x08;  // WORD
        const int OFF_FMT_WIDTH       = 0x0A;  // WORD
        const int OFF_FMT_FONTNUM     = 0x0C;  // WORD
        const int OFF_FMT_FONTSIZE    = 0x0E;  // WORD (знаковое: = -4*размер_пт)
        const int OFF_FMT_BOLD        = 0x10;
        const int OFF_FMT_ITALIC      = 0x11;
        const int OFF_FMT_UNDERLINE   = 0x12;
        const int OFF_FMT_HALIGN      = 0x13;
        const int OFF_FMT_VALIGN      = 0x14;
        const int OFF_FMT_PATTERN     = 0x15;
        const int OFF_FMT_BORDERL     = 0x16;
        const int OFF_FMT_BORDERT     = 0x17;
        const int OFF_FMT_BORDERR     = 0x18;
        const int OFF_FMT_BORDERB     = 0x19;
        const int OFF_FMT_PATTCOLOR   = 0x1A;
        const int OFF_FMT_BORDERCOLOR = 0x1B;
        const int OFF_FMT_TEXTCOLOR   = 0x1C;
        const int OFF_FMT_BKCOLOR     = 0x1D;
        const int OFF_FMT_WRAP        = 0x1E;
        const int OFF_FMT_DATAFORMAT  = 0x1F;
        const int OFF_FMT_UNPROTECT   = 0x20;
        const int OFF_FMT_TEXTDIR     = 0x22;  // WORD

        const int OFF_CELL_TEXT      = 0x24;
        const int OFF_CELL_STR2      = 0x28;
        const int OFF_CELL_STR3      = 0x2C;
        const int OFF_CELL_ARR_PDATA = 0x34;
        const int OFF_CELL_ARR_NSIZE = 0x38;

        // ------------------------- CSheetDrawing ----------------------------
        const int OFF_DWG_PCONTAINER = 0x44;
        const int OFF_DWG_PPICTURE   = 0x48;
        const int OFF_DWG_TYPE       = 0x4C;
        const int OFF_DWG_TL_XCELL   = 0x50;
        const int OFF_DWG_TL_YCELL   = 0x54;
        const int OFF_DWG_TL_XOFF    = 0x58;
        const int OFF_DWG_TL_YOFF    = 0x5C;
        const int OFF_DWG_BR_XCELL   = 0x60;
        const int OFF_DWG_BR_YCELL   = 0x64;
        const int OFF_DWG_BR_XOFF    = 0x68;
        const int OFF_DWG_BR_YOFF    = 0x6C;
        const int OFF_DWG_INDEX      = 0x70;

        // ---------------------------- LOGFONTA ------------------------------
        const int OFF_LF_HEIGHT    = 0x00;
        const int OFF_LF_WEIGHT    = 0x10;
        const int OFF_LF_ITALIC    = 0x14;
        const int OFF_LF_UNDERLINE = 0x15;
        const int OFF_LF_STRIKEOUT = 0x16;
        const int OFF_LF_CHARSET   = 0x17;
        const int OFF_LF_FACENAME  = 0x1C;
        const int SIZE_LOGFONTA    = 0x3C;

        // ------------------- маски свойств (MOXEL.H:104) --------------------
        public const int SFM_FONTNUMINCACHE  = 0x00000001;
        public const int SFM_FONTSIZE        = 0x00000002;
        public const int SFM_FONTBOLD        = 0x00000004;
        public const int SFM_FONTITALIC      = 0x00000008;
        public const int SFM_FONTUNDERLINE   = 0x00000010;
        public const int SFM_BORDERLEFT      = 0x00000020;   // == SFM_LINESTYLE
        public const int SFM_BORDERTOP       = 0x00000040;   // == SFM_LINEWEIGHT
        public const int SFM_BORDERRIGHT     = 0x00000080;   // == SFM_BORDERS
        public const int SFM_BORDERBOTTOM    = 0x00000100;   // == SFM_TYPEOUT
        public const int SFM_BORDERCOLOR     = 0x00000200;
        public const int SFM_HEIGHT          = 0x00000400;
        public const int SFM_WIDTH           = 0x00000800;
        public const int SFM_HALIGN          = 0x00001000;
        public const int SFM_VALIGN          = 0x00002000;
        public const int SFM_TEXTCOLOR       = 0x00004000;
        public const int SFM_BACKGROUNDCOLOR = 0x00008000;
        public const int SFM_PATTERN         = 0x00010000;
        public const int SFM_PATTERNCOLOR    = 0x00020000;
        public const int SFM_WRAPTEXT        = 0x00040000;
        public const int SFM_DATAFORMAT      = 0x00080000;
        public const int SFM_UNPROTECT       = 0x00100000;
        public const int SFM_UNK             = 0x00200000;   // == MoxelCellFlags.Data
        public const int SFM_TEXTDIRECTION   = 0x00400000;
        public const int SFM_DETAILS         = 0x40000000;   // == MoxelCellFlags.Value
        public const int SFM_TEXT            = unchecked((int)0x80000000); // == MoxelCellFlags.Text

        // ---------------------------- состояние -----------------------------

        readonly IntPtr _pSheet;       // CSheet*
        readonly IntPtr _pDoc;         // CSheetDoc*
        readonly object _comObject;    // RCW-держатель: пока жив он — жива таблица

        static readonly Encoding Cp1251 = Encoding.GetEncoding(1251);

        /// <summary>Из V7Table — основной путь.</summary>
        public MemSheet(V7Table table)
        {
            if (table == null)
                throw new ArgumentNullException("table");
            if (table.SheetDoc == null || table.SheetDoc.Pointer == IntPtr.Zero)
                throw new InvalidOperationException("MemSheet: таблица не установлена (V7Table.SetTable).");

            _comObject = table.ComObject;
            _pDoc = table.SheetDoc.Pointer;
            _pSheet = _pDoc + OFF_SHEETDOC_SHEET;
        }

        /// <summary>Напрямую по CSheetDoc* (держатель RCW — на вашей совести).</summary>
        public MemSheet(IntPtr pSheetDoc, object comObjectHolder)
        {
            if (pSheetDoc == IntPtr.Zero)
                throw new ArgumentException("pSheetDoc == IntPtr.Zero");
            _comObject = comObjectHolder;
            _pDoc = pSheetDoc;
            _pSheet = pSheetDoc + OFF_SHEETDOC_SHEET;
        }

        public IntPtr SheetPointer { get { return _pSheet; } }
        public IntPtr DocPointer { get { return _pDoc; } }

        /// <summary>Грубая проверка: память CSheet читаема и back-ptr сходится.</summary>
        public bool IsValid
        {
            get
            {
                if (!IsReadable(_pSheet, 0x220)) return false;
                IntPtr back = Marshal.ReadIntPtr(_pSheet, 0x21C);
                return back == IntPtr.Zero || back == _pDoc;
            }
        }

        // ====================================================================
        //  Размеры
        // ====================================================================

        /// <summary>m_nRows — логический размер таблицы (строк).</summary>
        public int RowsCount { get { return ReadI32(_pSheet, OFF_SHEET_NROWS); } }
        /// <summary>m_nCols — логический размер таблицы (колонок).</summary>
        public int ColsCount { get { return ReadI32(_pSheet, OFF_SHEET_NCOLS); } }

        /// <summary>Физических объектов-строк в m_RowsArray (разреженное хранение!).</summary>
        public int RowObjectsCount { get { return ReadI32(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_VAL_NSIZE); } }
        /// <summary>Форматов-колонок в m_ColumnsArray.</summary>
        public int ColFormatCount { get { return ReadI32(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_VAL_NSIZE); } }
        public int FontCount { get { return ReadI32(_pSheet, OFF_SHEET_FONTSARRAY + OFF_SORT_VAL_NSIZE); } }
        public int MaskCount { get { return ReadI32(_pSheet, OFF_SHEET_MASKSARRAY + OFF_SORT_VAL_NSIZE); } }
        public int DrawingCount { get { return ReadI32(_pSheet, OFF_SHEET_DRAWLIST_COUNT); } }

        // ====================================================================
        //  Колонтитулы (CSheetCell m_Header/m_Footer — встроены в CSheet)
        // ====================================================================

        public string HeaderText { get { return ReadCString(ReadPtr(_pSheet, OFF_SHEET_HEADER + OFF_CELL_TEXT)); } }
        public string FooterText { get { return ReadCString(ReadPtr(_pSheet, OFF_SHEET_FOOTER + OFF_CELL_TEXT)); } }

        /// <summary>Полные CSheetCell-структуры колонтитулов: Text + Str2/Str3 + формат.</summary>
        public SheetCellInfo GetHeaderCell()
        {
            var info = new SheetCellInfo { Row = -1, Col = -1, HasCell = true };
            FillCellStrings(info, _pSheet + OFF_SHEET_HEADER);
            ApplyFormat(info, _pSheet + OFF_SHEET_HEADER);
            ResolveFont(info);
            return info;
        }

        public SheetCellInfo GetFooterCell()
        {
            var info = new SheetCellInfo { Row = -1, Col = -1, HasCell = true };
            FillCellStrings(info, _pSheet + OFF_SHEET_FOOTER);
            ApplyFormat(info, _pSheet + OFF_SHEET_FOOTER);
            ResolveFont(info);
            return info;
        }

        // ====================================================================
        //  Поиск объектов строк/ячеек/колонок (разреженные keyed-массивы)
        // ====================================================================

        /// <summary>Найти CSheetRow* по номеру строки. false = у строки нет своих данных.</summary>
        public bool TryGetRowObject(int row, out IntPtr pRow)
        {
            IntPtr keys = ReadPtr(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_IDX_PDATA);
            int nKeys = ReadI32(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_IDX_NSIZE);
            IntPtr vals = ReadPtr(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_VAL_PDATA);
            int nVals = ReadI32(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_VAL_NSIZE);
            if (LookupKeyed(keys, nKeys, vals, nVals, row, out pRow))
                return pRow != IntPtr.Zero;
            return false;
        }

        /// <summary>Найти CSheetFormat* колонки. false = формат колонки не задан.</summary>
        public bool TryGetColFormat(int col, out IntPtr pFmt)
        {
            IntPtr keys = ReadPtr(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_IDX_PDATA);
            int nKeys = ReadI32(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_IDX_NSIZE);
            IntPtr vals = ReadPtr(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_VAL_PDATA);
            int nVals = ReadI32(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_VAL_NSIZE);
            if (LookupKeyed(keys, nKeys, vals, nVals, col, out pFmt))
                return pFmt != IntPtr.Zero;
            return false;
        }

        /// <summary>Найти CSheetCell* внутри конкретного CSheetRow.</summary>
        public bool TryGetCellAt(IntPtr pRow, int col, out IntPtr pCell)
        {
            pCell = IntPtr.Zero;
            if (pRow == IntPtr.Zero) return false;
            IntPtr keys = ReadPtr(pRow, OFF_ROW_KEYS_PDATA);
            int nKeys = ReadI32(pRow, OFF_ROW_KEYS_NSIZE);
            IntPtr vals = ReadPtr(pRow, OFF_ROW_CELLS_PDATA);
            int nVals = ReadI32(pRow, OFF_ROW_CELLS_NSIZE);
            if (LookupKeyed(keys, nKeys, vals, nVals, col, out pCell))
                return pCell != IntPtr.Zero;
            return false;
        }

        /// <summary>Найти CSheetCell* по (row, col).</summary>
        public bool TryGetCellObject(int row, int col, out IntPtr pCell)
        {
            pCell = IntPtr.Zero;
            IntPtr pRow;
            if (!TryGetRowObject(row, out pRow)) return false;
            return TryGetCellAt(pRow, col, out pCell);
        }

        /// <summary>
        /// Поиск ключа в параллельных массивах CSortArray/CSheetRow.
        /// Основной путь — бинарный поиск (массивы хранятся отсортированными),
        /// страховка — линейный проход (если порядок нарушен).
        /// </summary>
        static bool LookupKeyed(IntPtr pKeys, int nKeys, IntPtr pVals, int nVals,
                                int key, out IntPtr value)
        {
            value = IntPtr.Zero;
            int n = Math.Min(nKeys, nVals);
            if (n <= 0 || pKeys == IntPtr.Zero || pVals == IntPtr.Zero) return false;
            if (!IsReadable(pKeys, n * 4) || !IsReadable(pVals, n * 4)) return false;

            int lo = 0, hi = n - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                int k = Marshal.ReadInt32(pKeys, mid * 4);
                if (k == key)
                {
                    value = Marshal.ReadIntPtr(pVals, mid * 4);
                    return true;
                }
                if (k < key) lo = mid + 1; else hi = mid - 1;
            }

            // страховка от несортированности
            for (int i = 0; i < n; i++)
            {
                if (Marshal.ReadInt32(pKeys, i * 4) == key)
                {
                    value = Marshal.ReadIntPtr(pVals, i * 4);
                    return true;
                }
            }
            return false;
        }

        // ====================================================================
        //  Чтение ячеек
        // ====================================================================

        /// <summary>Индексатор: текст ячейки (0-based).</summary>
        public string this[int row, int col]
        {
            get { return GetCellText(row, col); }
        }

        /// <summary>Текст ячейки. Пустая строка, если объекта нет.</summary>
        public string GetCellText(int row, int col)
        {
            IntPtr pCell;
            if (!TryGetCellObject(row, col, out pCell)) return string.Empty;
            return ReadCString(ReadPtr(pCell, OFF_CELL_TEXT));
        }

        /// <summary>
        /// Ячейка с ЭФФЕКТИВНЫМ форматом: default листа ← колонка ← строка ← ячейка
        /// (аналог CSheet::GetCellAttributes; каждый слой накладывает только
        /// поля с выставленными битами маски).
        /// Никогда не возвращает null — при отсутствии объекта HasCell=false.
        /// </summary>
        public SheetCellInfo GetCell(int row, int col)
        {
            var info = new SheetCellInfo { Row = row, Col = col };

            // 1) default-формат листа (CSheet унаследован от CSheetFormat)
            ApplyFormat(info, _pSheet);

            // 2) формат колонки
            IntPtr pCol;
            if (TryGetColFormat(col, out pCol)) ApplyFormat(info, pCol);

            // 3) формат строки
            IntPtr pRow;
            if (TryGetRowObject(row, out pRow)) ApplyFormat(info, pRow);

            // 4) сама ячейка
            IntPtr pCell;
            if (pRow != IntPtr.Zero && TryGetCellAt(pRow, col, out pCell) && pCell != IntPtr.Zero)
            {
                info.HasCell = true;
                FillCellStrings(info, pCell);
                ApplyFormat(info, pCell);
                info.BytesCount = ReadI32(pCell, OFF_CELL_ARR_NSIZE);
            }

            ResolveFont(info);
            return info;
        }

        /// <summary>Формат строки (без наложения прочих слоёв).</summary>
        public SheetCellInfo GetRowFormat(int row)
        {
            var info = new SheetCellInfo { Row = row, Col = -1, Text = null };
            IntPtr pRow;
            if (TryGetRowObject(row, out pRow))
            {
                info.HasCell = false;
                ApplyFormat(info, pRow);
            }
            ResolveFont(info);
            return info;
        }

        /// <summary>Формат колонки (без наложения прочих слоёв).</summary>
        public SheetCellInfo GetColFormat(int col)
        {
            var info = new SheetCellInfo { Row = -1, Col = col, Text = null };
            IntPtr pFmt;
            if (TryGetColFormat(col, out pFmt))
                ApplyFormat(info, pFmt);
            ResolveFont(info);
            return info;
        }

        /// <summary>Текст/строки ячейки без изменения формата.</summary>
        static void FillCellStrings(SheetCellInfo info, IntPtr pCell)
        {
            info.Text = ReadCString(ReadPtr(pCell, OFF_CELL_TEXT));
            info.Str2 = ReadCString(ReadPtr(pCell, OFF_CELL_STR2));
            info.Str3 = ReadCString(ReadPtr(pCell, OFF_CELL_STR3));
        }

        void ResolveFont(SheetCellInfo info)
        {
            if (info.FontIndex.HasValue)
                info.FontName = GetFontName(info.FontIndex.Value);
        }

        /// <summary>
        /// Наложить формат (любой CSheetFormat*: лист/колонка/строка/ячейка):
        /// копируются только поля с выставленными битами маски.
        /// </summary>
        static void ApplyFormat(SheetCellInfo dst, IntPtr pFmt)
        {
            if (pFmt == IntPtr.Zero || !IsReadable(pFmt, 0x24)) return;
            int mask = Marshal.ReadInt32(pFmt, OFF_FMT_MASK);
            dst.Mask |= mask;

            if ((mask & SFM_FONTNUMINCACHE) != 0) dst.FontIndex = (int)ReadU16(pFmt, OFF_FMT_FONTNUM);
            if ((mask & SFM_FONTSIZE) != 0)
            {
                short raw = (short)ReadU16(pFmt, OFF_FMT_FONTSIZE);
                dst.FontSizeRaw = raw;
                dst.FontSizePt = raw / -4f;    // MOXEL.H: shdFontSize = -4*size
            }
            if ((mask & SFM_FONTBOLD) != 0)
            {
                byte b = ReadU8(pFmt, OFF_FMT_BOLD);
                dst.BoldRaw = b;
                dst.Bold = b == 0x07;          // clFontWeight: 0=Empty, 4=Normal, 7=Bold
            }
            if ((mask & SFM_FONTITALIC) != 0)
            {
                byte b = ReadU8(pFmt, OFF_FMT_ITALIC);
                dst.ItalicRaw = b; dst.Italic = b != 0;
            }
            if ((mask & SFM_FONTUNDERLINE) != 0)
            {
                byte b = ReadU8(pFmt, OFF_FMT_UNDERLINE);
                dst.UnderlineRaw = b; dst.Underline = b != 0;
            }

            if ((mask & SFM_HEIGHT) != 0) dst.Height = ReadU16(pFmt, OFF_FMT_HEIGHT);
            if ((mask & SFM_WIDTH) != 0) dst.Width = ReadU16(pFmt, OFF_FMT_WIDTH);
            if ((mask & SFM_HALIGN) != 0) dst.HAlign = ReadU8(pFmt, OFF_FMT_HALIGN);
            if ((mask & SFM_VALIGN) != 0) dst.VAlign = ReadU8(pFmt, OFF_FMT_VALIGN);

            if ((mask & SFM_BORDERLEFT) != 0) { dst.BorderLeft = ReadU8(pFmt, OFF_FMT_BORDERL); dst.LineStyle = dst.BorderLeft; }
            if ((mask & SFM_BORDERTOP) != 0) { dst.BorderTop = ReadU8(pFmt, OFF_FMT_BORDERT); dst.LineWeight = dst.BorderTop; }
            if ((mask & SFM_BORDERRIGHT) != 0) { dst.BorderRight = ReadU8(pFmt, OFF_FMT_BORDERR); dst.Borders = dst.BorderRight; }
            if ((mask & SFM_BORDERBOTTOM) != 0) { dst.BorderBottom = ReadU8(pFmt, OFF_FMT_BORDERB); dst.TypeOut = dst.BorderBottom; }

            if ((mask & SFM_BORDERCOLOR) != 0) dst.BorderColor = ReadU8(pFmt, OFF_FMT_BORDERCOLOR);
            if ((mask & SFM_TEXTCOLOR) != 0) dst.TextColor = ReadU8(pFmt, OFF_FMT_TEXTCOLOR);
            if ((mask & SFM_BACKGROUNDCOLOR) != 0) dst.BkColor = ReadU8(pFmt, OFF_FMT_BKCOLOR);
            if ((mask & SFM_PATTERN) != 0) dst.Pattern = ReadU8(pFmt, OFF_FMT_PATTERN);
            if ((mask & SFM_PATTERNCOLOR) != 0) dst.PatternColor = ReadU8(pFmt, OFF_FMT_PATTCOLOR);
            if ((mask & SFM_WRAPTEXT) != 0) dst.Wrap = ReadU8(pFmt, OFF_FMT_WRAP);
            if ((mask & SFM_DATAFORMAT) != 0) dst.DataFormat = ReadU8(pFmt, OFF_FMT_DATAFORMAT);
            if ((mask & SFM_UNPROTECT) != 0) dst.Unprotected = ReadU8(pFmt, OFF_FMT_UNPROTECT) != 0;
            if ((mask & SFM_TEXTDIRECTION) != 0) dst.TextDirection = ReadU16(pFmt, OFF_FMT_TEXTDIR);
            if ((mask & SFM_TEXT) != 0) dst.HasText = true;     // 0x80000000
            if ((mask & SFM_DETAILS) != 0) dst.HasValue = true; // 0x40000000
            if ((mask & SFM_UNK) != 0) dst.HasData = true;      // 0x00200000
        }

        // ====================================================================
        //  Массовое чтение
        // ====================================================================

        /// <summary>
        /// Прочитать ВСЕ существующие ячейки (обход разреженных массивов:
        /// m_RowsArray → CSheetRow → m_Cells). Формат — эффективный.
        /// </summary>
        public List<SheetCellInfo> ReadCells()
        {
            var list = new List<SheetCellInfo>();

            IntPtr rowKeys = ReadPtr(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_IDX_PDATA);
            int nRowKeys = ReadI32(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_IDX_NSIZE);
            IntPtr rowVals = ReadPtr(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_VAL_PDATA);
            int nRowVals = ReadI32(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_VAL_NSIZE);

            int n = Math.Min(nRowKeys, nRowVals);
            if (n <= 0 || rowKeys == IntPtr.Zero || rowVals == IntPtr.Zero) return list;
            if (!IsReadable(rowKeys, n * 4) || !IsReadable(rowVals, n * 4)) return list;

            for (int i = 0; i < n; i++)
            {
                int rowKey = Marshal.ReadInt32(rowKeys, i * 4);
                IntPtr pRow = Marshal.ReadIntPtr(rowVals, i * 4);
                if (pRow == IntPtr.Zero || !IsReadable(pRow, 0x50)) continue;

                IntPtr cellKeys = ReadPtr(pRow, OFF_ROW_KEYS_PDATA);
                int nCellKeys = ReadI32(pRow, OFF_ROW_KEYS_NSIZE);
                IntPtr cellVals = ReadPtr(pRow, OFF_ROW_CELLS_PDATA);
                int nCellVals = ReadI32(pRow, OFF_ROW_CELLS_NSIZE);

                int m = Math.Min(nCellKeys, nCellVals);
                if (m <= 0 || cellKeys == IntPtr.Zero || cellVals == IntPtr.Zero) continue;
                if (!IsReadable(cellKeys, m * 4) || !IsReadable(cellVals, m * 4)) continue;

                for (int j = 0; j < m; j++)
                {
                    int colKey = Marshal.ReadInt32(cellKeys, j * 4);
                    IntPtr pCell = Marshal.ReadIntPtr(cellVals, j * 4);
                    if (pCell == IntPtr.Zero) continue;
                    list.Add(GetCell(rowKey, colKey));
                }
            }
            return list;
        }

        /// <summary>
        /// Плотная сетка текстов rows×cols (0-based). maxRows/maxCols=0 — без
        /// ограничения (осторожно с большими таблицами!).
        /// </summary>
        public string[,] ReadTextGrid(int maxRows, int maxCols)
        {
            int rows = RowsCount, cols = ColsCount;
            if (maxRows > 0) rows = Math.Min(rows, maxRows);
            if (maxCols > 0) cols = Math.Min(cols, maxCols);
            if (rows < 0) rows = 0;
            if (cols < 0) cols = 0;

            var grid = new string[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    grid[r, c] = GetCellText(r, c);
            return grid;
        }

        // ====================================================================
        //  Шрифты (CSheetFontsArray: LOGFONTA по значению, 0x3C на элемент)
        // ====================================================================

        /// <summary>Имя шрифта по индексу кэша шрифтов. null — вне диапазона.</summary>
        public string GetFontName(int index)
        {
            IntPtr pData = ReadPtr(_pSheet, OFF_SHEET_FONTSARRAY + OFF_SORT_VAL_PDATA);
            int n = ReadI32(_pSheet, OFF_SHEET_FONTSARRAY + OFF_SORT_VAL_NSIZE);
            if (pData == IntPtr.Zero || index < 0 || index >= n) return null;
            IntPtr lf = pData + index * SIZE_LOGFONTA;
            if (!IsReadable(lf, SIZE_LOGFONTA)) return null;
            return ReadAnsiFixed(lf + OFF_LF_FACENAME, 32);
        }

        /// <summary>Краткое описание шрифта (для дампа/отладки).</summary>
        public string GetFontSummary(int index)
        {
            IntPtr pData = ReadPtr(_pSheet, OFF_SHEET_FONTSARRAY + OFF_SORT_VAL_PDATA);
            int n = ReadI32(_pSheet, OFF_SHEET_FONTSARRAY + OFF_SORT_VAL_NSIZE);
            if (pData == IntPtr.Zero || index < 0 || index >= n) return "<вне диапазона>";
            IntPtr lf = pData + index * SIZE_LOGFONTA;
            if (!IsReadable(lf, SIZE_LOGFONTA)) return "<нечитаемо>";

            string name = ReadAnsiFixed(lf + OFF_LF_FACENAME, 32);
            int h = Marshal.ReadInt32(lf, OFF_LF_HEIGHT);
            int w = Marshal.ReadInt32(lf, OFF_LF_WEIGHT);
            bool it = Marshal.ReadByte(lf, OFF_LF_ITALIC) != 0;
            bool un = Marshal.ReadByte(lf, OFF_LF_UNDERLINE) != 0;
            bool so = Marshal.ReadByte(lf, OFF_LF_STRIKEOUT) != 0;
            return string.Format("'{0}' h={1} weight={2}{3}{4}{5}",
                name, h, w, it ? " italic" : "", un ? " underline" : "", so ? " strikeout" : "");
        }

        /// <summary>Маска формата данных (CSheetMasksArray) по индексу.</summary>
        public string GetMask(int index)
        {
            IntPtr pData = ReadPtr(_pSheet, OFF_SHEET_MASKSARRAY + OFF_SORT_VAL_PDATA);
            int n = ReadI32(_pSheet, OFF_SHEET_MASKSARRAY + OFF_SORT_VAL_NSIZE);
            if (pData == IntPtr.Zero || index < 0 || index >= n) return null;
            if (!IsReadable(pData + index * 4, 4)) return null;
            return ReadCString(Marshal.ReadIntPtr(pData, index * 4));
        }

        // ====================================================================
        //  Рисунки (CSheetDrawingList) и секции (CSheetOutlineArray)
        // ====================================================================

        public List<SheetDrawingInfo> ReadDrawings()
        {
            var res = new List<SheetDrawingInfo>();
            IntPtr node = ReadPtr(_pSheet, OFF_SHEET_DRAWLIST_HEAD);
            int count = ReadI32(_pSheet, OFF_SHEET_DRAWLIST_COUNT);
            if (node == IntPtr.Zero || count <= 0) return res;

            for (int i = 0; i < count && node != IntPtr.Zero; i++)
            {
                IntPtr pDwg = ReadPtr(node, 8);   // CList::CNode {pNext, pPrev, data}
                if (pDwg != IntPtr.Zero && IsReadable(pDwg, 0x74))
                {
                    var d = new SheetDrawingInfo();
                    d.Index = ReadI32(pDwg, OFF_DWG_INDEX);
                    d.Type = ReadI32(pDwg, OFF_DWG_TYPE);
                    d.Col1 = ReadI32(pDwg, OFF_DWG_TL_XCELL);
                    d.Row1 = ReadI32(pDwg, OFF_DWG_TL_YCELL);
                    d.XOff1 = ReadI32(pDwg, OFF_DWG_TL_XOFF);
                    d.YOff1 = ReadI32(pDwg, OFF_DWG_TL_YOFF);
                    d.Col2 = ReadI32(pDwg, OFF_DWG_BR_XCELL);
                    d.Row2 = ReadI32(pDwg, OFF_DWG_BR_YCELL);
                    d.XOff2 = ReadI32(pDwg, OFF_DWG_BR_XOFF);
                    d.YOff2 = ReadI32(pDwg, OFF_DWG_BR_YOFF);
                    d.pPicture = ReadPtr(pDwg, OFF_DWG_PPICTURE);
                    d.pContainer = ReadPtr(pDwg, OFF_DWG_PCONTAINER);
                    d.Text = ReadCString(ReadPtr(pDwg, OFF_CELL_TEXT));
                    d.Str2 = ReadCString(ReadPtr(pDwg, OFF_CELL_STR2));
                    res.Add(d);
                }
                node = ReadPtr(node, 0);          // pNext
            }
            return res;
        }

        /// <summary>Секции (группировки). vertical=false — по строкам, true — по колонкам.</summary>
        public List<SheetSectionInfo> ReadSections(bool vertical)
        {
            var res = new List<SheetSectionInfo>();
            int off = vertical ? OFF_SHEET_VERTSECTIONS : OFF_SHEET_HORZSECTIONS;

            // CArray<CSheetOutline>: vft@0, pData@+4, nSize@+8
            IntPtr pData = ReadPtr(_pSheet, off + 4);
            int n = ReadI32(_pSheet, off + 8);
            if (pData == IntPtr.Zero || n <= 0 || n > 65536) return res;
            if (!IsReadable(pData, n * 0x14)) return res;

            for (int i = 0; i < n; i++)
            {
                IntPtr e = pData + i * 0x14;      // CSheetOutline{vft,Start,End,data,Name}
                var s = new SheetSectionInfo { Vertical = vertical };
                s.Start = Marshal.ReadInt32(e, 4);
                s.End = Marshal.ReadInt32(e, 8);
                s.Name = ReadCString(Marshal.ReadIntPtr(e, 0x10));
                res.Add(s);
            }
            return res;
        }

        // ====================================================================
        //  Гипотезы (печатаются в Dump() для калибровки на живой 1С)
        // ====================================================================

        /// <summary>
        /// Разрывы страниц — ГИПОТЕЗА: в зоне 0x184..0x1AC ровно встают
        /// два CArray&lt;int&gt; (vft+pData+nSize+nMax+nGrowBy @0x184 и @0x198,
        /// конец 0x1AC — начало списка рисунков). В MXL-файле разрывы пишутся
        /// после секций (V, затем H). index = 0 или 1; какой из них
        /// горизонтальный/вертикальный — сверить на живой таблице.
        /// </summary>
        public int[] ReadPageBreaks(int index)
        {
            int offData = index == 0 ? OFF_SHEET_PB1_PDATA : OFF_SHEET_PB2_PDATA;
            int offSize = index == 0 ? OFF_SHEET_PB1_NSIZE : OFF_SHEET_PB2_NSIZE;

            IntPtr pData = ReadPtr(_pSheet, offData);
            int n = ReadI32(_pSheet, offSize);
            if (pData == IntPtr.Zero || n <= 0 || n > 10000) return null;
            if (!IsReadable(pData, n * 4)) return null;

            var res = new int[n];
            for (int i = 0; i < n; i++)
                res[i] = Marshal.ReadInt32(pData, i * 4);
            return res;
        }

        // ====================================================================
        //  Диагностика
        // ====================================================================

        /// <summary>Полный дамп содержимого листа (для калибровки смещений).</summary>
        public string Dump(int maxRows = 200)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Format("=== MemSheet @ 0x{0:X8} (CSheetDoc 0x{1:X8}) valid={2} ===",
                _pSheet.ToInt32(), _pDoc.ToInt32(), IsValid));
            if (!IsReadable(_pSheet, 0x220))
            {
                sb.AppendLine("Память CSheet недоступна!");
                return sb.ToString();
            }

            sb.AppendLine(string.Format("Размер: m_nRows={0}, m_nCols={1}; объектов-строк={2}, форматов-колонок={3}, шрифтов={4}, масок={5}, рисунков={6}",
                RowsCount, ColsCount, RowObjectsCount, ColFormatCount, FontCount, MaskCount, DrawingCount));

            sb.AppendLine(string.Format("Колонтитул верх: '{0}'", Clip(HeaderText, 100)));
            sb.AppendLine(string.Format("Колонтитул низ : '{0}'", Clip(FooterText, 100)));

            int fc = Math.Min(FontCount, 32);
            for (int i = 0; i < fc; i++)
                sb.AppendLine(string.Format("  font[{0}] = {1}", i, GetFontSummary(i)));

            int mc = Math.Min(MaskCount, 16);
            for (int i = 0; i < mc; i++)
                sb.AppendLine(string.Format("  mask[{0}] = '{1}'", i, Clip(GetMask(i), 48)));

            // --- колонки ---
            IntPtr ck = ReadPtr(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_IDX_PDATA);
            int ckn = ReadI32(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_IDX_NSIZE);
            IntPtr cv = ReadPtr(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_VAL_PDATA);
            int cvn = ReadI32(_pSheet, OFF_SHEET_COLUMNSARRAY + OFF_SORT_VAL_NSIZE);
            int cn = Math.Min(Math.Min(ckn, cvn), 48);
            if (cn > 0 && ck != IntPtr.Zero && cv != IntPtr.Zero)
            {
                sb.AppendLine(string.Format("Форматы колонок ({0}):", ColFormatCount));
                for (int i = 0; i < cn; i++)
                {
                    int key = Marshal.ReadInt32(ck, i * 4);
                    IntPtr pFmt = Marshal.ReadIntPtr(cv, i * 4);
                    if (pFmt == IntPtr.Zero || !IsReadable(pFmt, 0x24)) continue;
                    sb.AppendLine(string.Format("  col[{0}]: width={1} mask=0x{2:X8}",
                        key, ReadU16(pFmt, OFF_FMT_WIDTH), Marshal.ReadInt32(pFmt, OFF_FMT_MASK)));
                }
            }

            // --- строки и ячейки ---
            IntPtr rk = ReadPtr(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_IDX_PDATA);
            int rkn = ReadI32(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_IDX_NSIZE);
            IntPtr rv = ReadPtr(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_VAL_PDATA);
            int rvn = ReadI32(_pSheet, OFF_SHEET_ROWSARRAY + OFF_SORT_VAL_NSIZE);
            int totalRows = Math.Min(rkn, rvn);
            int rn = Math.Min(totalRows, maxRows);
            if (rn > 0 && rk != IntPtr.Zero && rv != IntPtr.Zero)
            {
                sb.AppendLine(string.Format("Строки ({0} объектов):", totalRows));
                for (int i = 0; i < rn; i++)
                {
                    int rowKey = Marshal.ReadInt32(rk, i * 4);
                    IntPtr pRow = Marshal.ReadIntPtr(rv, i * 4);
                    if (pRow == IntPtr.Zero || !IsReadable(pRow, 0x50)) continue;

                    sb.AppendLine(string.Format("  row[{0}]: height={1} mask=0x{2:X8} rtti='{3}'",
                        rowKey, ReadU16(pRow, OFF_FMT_HEIGHT),
                        Marshal.ReadInt32(pRow, OFF_FMT_MASK),
                        V7Table.TryGetRuntimeClassName(pRow)));

                    IntPtr cKeys = ReadPtr(pRow, OFF_ROW_KEYS_PDATA);
                    int ckn2 = ReadI32(pRow, OFF_ROW_KEYS_NSIZE);
                    IntPtr cVals = ReadPtr(pRow, OFF_ROW_CELLS_PDATA);
                    int cvn2 = ReadI32(pRow, OFF_ROW_CELLS_NSIZE);
                    int m = Math.Min(ckn2, cvn2);
                    if (m > 0 && cKeys != IntPtr.Zero && cVals != IntPtr.Zero)
                    {
                        for (int j = 0; j < m; j++)
                        {
                            int colKey = Marshal.ReadInt32(cKeys, j * 4);
                            IntPtr pCell = Marshal.ReadIntPtr(cVals, j * 4);
                            if (pCell == IntPtr.Zero || !IsReadable(pCell, 0x44)) continue;

                            string t = ReadCString(ReadPtr(pCell, OFF_CELL_TEXT));
                            string s2 = ReadCString(ReadPtr(pCell, OFF_CELL_STR2));
                            string s3 = ReadCString(ReadPtr(pCell, OFF_CELL_STR3));
                            sb.AppendLine(string.Format(
                                "    cell[{0}]: '{1}' | str2='{2}' | str3='{3}' | mask=0x{4:X8} font={5} size={6} arr={7}",
                                colKey, Clip(t, 60), Clip(s2, 32), Clip(s3, 32),
                                Marshal.ReadInt32(pCell, OFF_FMT_MASK),
                                ReadU16(pCell, OFF_FMT_FONTNUM),
                                (short)ReadU16(pCell, OFF_FMT_FONTSIZE),
                                ReadI32(pCell, OFF_CELL_ARR_NSIZE)));
                        }
                    }
                }
                if (rn < totalRows)
                    sb.AppendLine(string.Format("  ... (показано {0} из {1})", rn, totalRows));
            }

            // --- секции ---
            foreach (bool vert in new[] { false, true })
            {
                var secs = ReadSections(vert);
                if (secs == null || secs.Count == 0) continue;
                sb.AppendLine(string.Format("Секции {0}: {1}",
                    vert ? "вертикальные" : "горизонтальные", secs.Count));
                foreach (var s in secs)
                    sb.AppendLine(string.Format("  [{0}..{1}] '{2}'", s.Start, s.End, s.Name));
            }

            // --- рисунки ---
            var dwgs = ReadDrawings();
            if (dwgs != null && dwgs.Count > 0)
            {
                sb.AppendLine(string.Format("Рисунки: {0}", dwgs.Count));
                foreach (var d in dwgs)
                    sb.AppendLine("  " + d);
            }

            // --- гипотезы ---
            int[] pb1 = ReadPageBreaks(0);
            int[] pb2 = ReadPageBreaks(1);
            sb.AppendLine(string.Format("Разрывы страниц (два CArray<int> @0x184/0x198): PB1={0}, PB2={1} (какой H/V — сверить с 1С)",
                pb1 == null ? "<не массив>" : JoinInts(pb1, 24),
                pb2 == null ? "<не массив>" : JoinInts(pb2, 24)));

            // UnkArr2/3 — кандидаты на список объединений ячеек: элемент CPtrArray
            // = CRect* (16 байт left,top,right,bottom) = байт-в-байт CellsUnion
            sb.AppendLine("Кандидаты на объединения ячеек (CPtrArray → CRect*):");
            DumpRectArray(sb, "  UnkArr2 @0x1F4", 0x1F8, 0x1FC);
            DumpRectArray(sb, "  UnkArr3 @0x208", 0x20C, 0x210);

            return sb.ToString();
        }

        /// <summary>Дамп в Debug/Trace.</summary>
        public void DumpToDebug(int maxRows = 200)
        {
            string s = Dump(maxRows);
            foreach (string line in s.Split('\n'))
                Debug.WriteLine(line.TrimEnd('\r'));
        }

        static string Clip(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace("\r\n", "\\n").Replace("\r", "\\n")
                 .Replace("\n", "\\n").Replace("\t", "\\t");
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }

        static string JoinInts(int[] arr, int max)
        {
            var sb = new StringBuilder("[");
            for (int i = 0; i < arr.Length && i < max; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append(arr[i]);
            }
            if (arr.Length > max) sb.Append(",...");
            return sb.Append(']').ToString();
        }

        static string DumpDwords(IntPtr p, int off, int count)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                sb.Append(i == 0 ? "" : " ");
                sb.Append(string.Format("0x{0:X8}", ReadI32(p, off + i * 4)));
            }
            return sb.ToString();
        }

        /// <summary>Дамп CPtrArray с элементами-указателями на CRect (16 байт) —
        /// проверка гипотезы «объединения ячеек» (CellsUnion).</summary>
        void DumpRectArray(StringBuilder sb, string title, int offData, int offSize)
        {
            IntPtr pData = ReadPtr(_pSheet, offData);
            int n = ReadI32(_pSheet, offSize);
            if (pData == IntPtr.Zero || n <= 0)
            {
                sb.AppendLine(string.Format("{0}: пусто (n={1})", title, n));
                return;
            }

            int m = Math.Min(n, 24);
            var parts = new List<string>();
            for (int i = 0; i < m; i++)
            {
                IntPtr p = ReadPtr(pData, i * 4);
                if (p == IntPtr.Zero || !IsReadable(p, 16)) { parts.Add("<bad>"); continue; }
                parts.Add(string.Format("({0},{1})-({2},{3})",
                    Marshal.ReadInt32(p, 0), Marshal.ReadInt32(p, 4),
                    Marshal.ReadInt32(p, 8), Marshal.ReadInt32(p, 12)));
            }
            sb.AppendLine(string.Format("{0}: n={1} {2}{3}", title, n,
                string.Join(" ", parts.ToArray()), n > m ? " ..." : ""));
        }

        // ====================================================================
        //  CString (MFC42/VC6) и безопасные чтения
        // ====================================================================

        /// <summary>
        /// Прочитать MFC42 CString по указателю на СИМВОЛЬНЫЕ данные (m_pchData).
        /// Заголовок CStringData {nRefs; nDataLength; nAllocLength;} лежит
        /// непосредственно перед данными → длина = *(int*)(pch-8).
        /// Пустые CString указывают на статический _afxPchNil в mfc42.dll —
        /// заголовок там читаем, длина = 0.
        /// </summary>
        internal static string ReadCString(IntPtr pch)
        {
            if (pch == IntPtr.Zero) return string.Empty;

            if (IsReadable(pch - 12, 12))
            {
                int len = Marshal.ReadInt32(pch, -8);
                if (len == 0) return string.Empty;
                if (len > 0 && len < 0x100000 && IsReadable(pch, len))
                {
                    var buf = new byte[len];
                    Marshal.Copy(pch, buf, 0, len);
                    return Cp1251.GetString(buf);
                }
            }
            return ReadAnsiCapped(pch, 4096);
        }

        static string ReadAnsiCapped(IntPtr p, int cap)
        {
            var list = new List<byte>(64);
            for (int i = 0; i < cap; i++)
            {
                if (!IsReadable(p + i, 1)) break;
                byte b = Marshal.ReadByte(p, i);
                if (b == 0) break;
                list.Add(b);
            }
            return list.Count == 0 ? string.Empty : Cp1251.GetString(list.ToArray());
        }

        static string ReadAnsiFixed(IntPtr p, int max)
        {
            if (!IsReadable(p, max)) return string.Empty;
            var buf = new byte[max];
            Marshal.Copy(p, buf, 0, max);
            int len = 0;
            while (len < max && buf[len] != 0) len++;
            return Cp1251.GetString(buf, 0, len);
        }

        // ------------------------------------------------------------------
        //  Безопасные чтения (каждое — с проверкой IsBadReadPtr)
        // ------------------------------------------------------------------

        static IntPtr ReadPtr(IntPtr p, int off)
        {
            if (!IsReadable(p + off, 4)) return IntPtr.Zero;
            return Marshal.ReadIntPtr(p, off);
        }

        static int ReadI32(IntPtr p, int off)
        {
            if (!IsReadable(p + off, 4)) return 0;
            return Marshal.ReadInt32(p, off);
        }

        static ushort ReadU16(IntPtr p, int off)
        {
            if (!IsReadable(p + off, 2)) return 0;
            return (ushort)Marshal.ReadInt16(p, off);
        }

        static byte ReadU8(IntPtr p, int off)
        {
            if (!IsReadable(p + off, 1)) return 0;
            return Marshal.ReadByte(p, off);
        }

        internal static bool IsReadable(IntPtr p, int size)
        {
            if (p == IntPtr.Zero || size <= 0) return false;
            int addr = p.ToInt32();
            if (addr < 0x00010000 || addr >= 0x7FFF0000) return false;
            return IsBadReadPtr(p, (uint)size) == 0;
        }

        [DllImport("kernel32.dll")]
        static extern int IsBadReadPtr(IntPtr lp, uint ucb);
    }
}
