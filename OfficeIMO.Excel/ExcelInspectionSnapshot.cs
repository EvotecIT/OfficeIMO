namespace OfficeIMO.Excel {
    /// <summary>
    /// Immutable workbook inspection snapshot exposed by OfficeIMO.Excel for downstream integrations.
    /// </summary>
    public sealed class ExcelWorkbookSnapshot {
        private readonly List<ExcelWorksheetSnapshot> _worksheets = new List<ExcelWorksheetSnapshot>();
        private readonly List<ExcelNamedRangeSnapshot> _namedRanges = new List<ExcelNamedRangeSnapshot>();

        /// <summary>
        /// Workbook title, when present in package properties.
        /// </summary>
        public string? Title { get; internal set; }

        /// <summary>
        /// File path associated with the workbook, when the document was created from a path.
        /// </summary>
        public string? FilePath { get; internal set; }

        /// <summary>
        /// Worksheets in workbook order.
        /// </summary>
        public IReadOnlyList<ExcelWorksheetSnapshot> Worksheets => _worksheets;

        /// <summary>
        /// Workbook and sheet-local named ranges discovered during inspection.
        /// </summary>
        public IReadOnlyList<ExcelNamedRangeSnapshot> NamedRanges => _namedRanges;

        internal void AddWorksheet(ExcelWorksheetSnapshot worksheet) {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            _worksheets.Add(worksheet);
        }

        internal void AddNamedRange(ExcelNamedRangeSnapshot namedRange) {
            if (namedRange == null) throw new ArgumentNullException(nameof(namedRange));
            _namedRanges.Add(namedRange);
        }
    }

    /// <summary>
    /// Immutable worksheet inspection snapshot.
    /// </summary>
    public sealed class ExcelWorksheetSnapshot {
        private readonly List<ExcelCellSnapshot> _cells = new List<ExcelCellSnapshot>();
        private readonly List<ExcelMergedRangeSnapshot> _mergedRanges = new List<ExcelMergedRangeSnapshot>();
        private readonly List<ExcelColumnSnapshot> _columns = new List<ExcelColumnSnapshot>();
        private readonly List<ExcelRowSnapshot> _rows = new List<ExcelRowSnapshot>();
        private readonly List<ExcelTableSnapshot> _tables = new List<ExcelTableSnapshot>();

        /// <summary>
        /// Worksheet name.
        /// </summary>
        public string Name { get; internal set; } = string.Empty;

        /// <summary>
        /// Zero-based worksheet index in workbook order.
        /// </summary>
        public int Index { get; internal set; }

        /// <summary>
        /// Whether the worksheet is hidden.
        /// </summary>
        public bool Hidden { get; internal set; }

        /// <summary>
        /// Whether the worksheet is displayed right-to-left.
        /// </summary>
        public bool RightToLeft { get; internal set; }

        /// <summary>
        /// Worksheet tab color in ARGB hexadecimal form, when present.
        /// </summary>
        public string? TabColorArgb { get; internal set; }

        /// <summary>
        /// Number of frozen rows detected on the worksheet.
        /// </summary>
        public int FrozenRowCount { get; internal set; }

        /// <summary>
        /// Number of frozen columns detected on the worksheet.
        /// </summary>
        public int FrozenColumnCount { get; internal set; }

        /// <summary>
        /// Used-range address in A1 notation.
        /// </summary>
        public string UsedRangeA1 { get; internal set; } = "A1:A1";

        /// <summary>
        /// Non-empty cells discovered during inspection.
        /// </summary>
        public IReadOnlyList<ExcelCellSnapshot> Cells => _cells;

        /// <summary>
        /// Merged ranges discovered during inspection.
        /// </summary>
        public IReadOnlyList<ExcelMergedRangeSnapshot> MergedRanges => _mergedRanges;

        /// <summary>
        /// Explicit column definitions discovered during inspection.
        /// </summary>
        public IReadOnlyList<ExcelColumnSnapshot> Columns => _columns;

        /// <summary>
        /// Explicit row definitions discovered during inspection.
        /// </summary>
        public IReadOnlyList<ExcelRowSnapshot> Rows => _rows;

        /// <summary>
        /// Worksheet-level auto filter discovered during inspection, when present.
        /// </summary>
        public ExcelAutoFilterSnapshot? AutoFilter { get; internal set; }

        /// <summary>
        /// Worksheet protection discovered during inspection, when present.
        /// </summary>
        public ExcelWorksheetProtectionSnapshot? Protection { get; internal set; }

        /// <summary>
        /// Tables discovered on the worksheet.
        /// </summary>
        public IReadOnlyList<ExcelTableSnapshot> Tables => _tables;

        internal void AddCell(ExcelCellSnapshot cell) {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            _cells.Add(cell);
        }

        internal void AddMergedRange(ExcelMergedRangeSnapshot mergedRange) {
            if (mergedRange == null) throw new ArgumentNullException(nameof(mergedRange));
            _mergedRanges.Add(mergedRange);
        }

        internal void AddColumn(ExcelColumnSnapshot column) {
            if (column == null) throw new ArgumentNullException(nameof(column));
            _columns.Add(column);
        }

        internal void AddRow(ExcelRowSnapshot row) {
            if (row == null) throw new ArgumentNullException(nameof(row));
            _rows.Add(row);
        }

        internal void AddTable(ExcelTableSnapshot table) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            _tables.Add(table);
        }
    }

    /// <summary>
    /// Immutable worksheet-protection metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelWorksheetProtectionSnapshot {
        /// <summary>
        /// Whether selecting locked cells is allowed.
        /// </summary>
        public bool AllowSelectLockedCells { get; internal set; }

        /// <summary>
        /// Whether selecting unlocked cells is allowed.
        /// </summary>
        public bool AllowSelectUnlockedCells { get; internal set; }

        /// <summary>
        /// Whether formatting cells is allowed.
        /// </summary>
        public bool AllowFormatCells { get; internal set; }

        /// <summary>
        /// Whether formatting columns is allowed.
        /// </summary>
        public bool AllowFormatColumns { get; internal set; }

        /// <summary>
        /// Whether formatting rows is allowed.
        /// </summary>
        public bool AllowFormatRows { get; internal set; }

        /// <summary>
        /// Whether inserting columns is allowed.
        /// </summary>
        public bool AllowInsertColumns { get; internal set; }

        /// <summary>
        /// Whether inserting rows is allowed.
        /// </summary>
        public bool AllowInsertRows { get; internal set; }

        /// <summary>
        /// Whether inserting hyperlinks is allowed.
        /// </summary>
        public bool AllowInsertHyperlinks { get; internal set; }

        /// <summary>
        /// Whether deleting columns is allowed.
        /// </summary>
        public bool AllowDeleteColumns { get; internal set; }

        /// <summary>
        /// Whether deleting rows is allowed.
        /// </summary>
        public bool AllowDeleteRows { get; internal set; }

        /// <summary>
        /// Whether sorting is allowed.
        /// </summary>
        public bool AllowSort { get; internal set; }

        /// <summary>
        /// Whether AutoFilter interaction is allowed.
        /// </summary>
        public bool AllowAutoFilter { get; internal set; }

        /// <summary>
        /// Whether PivotTable interaction is allowed.
        /// </summary>
        public bool AllowPivotTables { get; internal set; }
    }

    /// <summary>
    /// Immutable cell inspection snapshot.
    /// </summary>
    public sealed class ExcelCellSnapshot {
        /// <summary>
        /// One-based row index.
        /// </summary>
        public int Row { get; internal set; }

        /// <summary>
        /// One-based column index.
        /// </summary>
        public int Column { get; internal set; }

        /// <summary>
        /// Typed cell value as interpreted by OfficeIMO's read model.
        /// </summary>
        public object? Value { get; internal set; }

        /// <summary>
        /// Formula text without a guaranteed leading equals sign.
        /// </summary>
        public string? Formula { get; internal set; }

        /// <summary>
        /// OpenXML style index currently backing the cell, when present.
        /// </summary>
        public uint? StyleIndex { get; internal set; }

        /// <summary>
        /// Resolved OfficeIMO style metadata for the cell, when available.
        /// </summary>
        public ExcelCellStyleSnapshot? Style { get; internal set; }

        /// <summary>
        /// Hyperlink metadata attached to the cell, when present.
        /// </summary>
        public ExcelHyperlinkSnapshot? Hyperlink { get; internal set; }

        /// <summary>
        /// Comment metadata attached to the cell, when present.
        /// </summary>
        public ExcelCommentSnapshot? Comment { get; internal set; }
    }

    /// <summary>
    /// Immutable comment metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelCommentSnapshot {
        /// <summary>
        /// Comment author display name, when available.
        /// </summary>
        public string? Author { get; internal set; }

        /// <summary>
        /// Comment text content.
        /// </summary>
        public string Text { get; internal set; } = string.Empty;
    }

    /// <summary>
    /// Immutable merge-range inspection snapshot.
    /// </summary>
    public sealed class ExcelMergedRangeSnapshot {
        /// <summary>
        /// Merged range in A1 notation.
        /// </summary>
        public string A1Range { get; internal set; } = string.Empty;

        /// <summary>
        /// One-based starting row of the merge.
        /// </summary>
        public int StartRow { get; internal set; }

        /// <summary>
        /// One-based ending row of the merge.
        /// </summary>
        public int EndRow { get; internal set; }

        /// <summary>
        /// One-based starting column of the merge.
        /// </summary>
        public int StartColumn { get; internal set; }

        /// <summary>
        /// One-based ending column of the merge.
        /// </summary>
        public int EndColumn { get; internal set; }
    }

    /// <summary>
    /// Immutable explicit column metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelColumnSnapshot {
        /// <summary>
        /// One-based starting column index covered by the definition.
        /// </summary>
        public int StartIndex { get; internal set; }

        /// <summary>
        /// One-based ending column index covered by the definition.
        /// </summary>
        public int EndIndex { get; internal set; }

        /// <summary>
        /// Excel column width in character units, when explicitly set.
        /// </summary>
        public double? Width { get; internal set; }

        /// <summary>
        /// Whether the definition represents a hidden column range.
        /// </summary>
        public bool Hidden { get; internal set; }

        /// <summary>
        /// Whether the width was explicitly customized.
        /// </summary>
        public bool CustomWidth { get; internal set; }
    }

    /// <summary>
    /// Immutable explicit row metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelRowSnapshot {
        /// <summary>
        /// One-based row index.
        /// </summary>
        public int Index { get; internal set; }

        /// <summary>
        /// Row height in points, when explicitly set.
        /// </summary>
        public double? Height { get; internal set; }

        /// <summary>
        /// Whether the row is hidden.
        /// </summary>
        public bool Hidden { get; internal set; }

        /// <summary>
        /// Whether the height was explicitly customized.
        /// </summary>
        public bool CustomHeight { get; internal set; }
    }

    /// <summary>
    /// Immutable auto-filter metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelAutoFilterSnapshot {
        private readonly List<ExcelFilterColumnSnapshot> _columns = new List<ExcelFilterColumnSnapshot>();

        /// <summary>
        /// Filtered range in A1 notation.
        /// </summary>
        public string A1Range { get; internal set; } = string.Empty;

        /// <summary>
        /// One-based starting row of the filter range.
        /// </summary>
        public int StartRow { get; internal set; }

        /// <summary>
        /// One-based ending row of the filter range.
        /// </summary>
        public int EndRow { get; internal set; }

        /// <summary>
        /// One-based starting column of the filter range.
        /// </summary>
        public int StartColumn { get; internal set; }

        /// <summary>
        /// One-based ending column of the filter range.
        /// </summary>
        public int EndColumn { get; internal set; }

        /// <summary>
        /// Filter criteria defined per relative column.
        /// </summary>
        public IReadOnlyList<ExcelFilterColumnSnapshot> Columns => _columns;

        internal void AddColumn(ExcelFilterColumnSnapshot column) {
            if (column == null) throw new ArgumentNullException(nameof(column));
            _columns.Add(column);
        }
    }

    /// <summary>
    /// Immutable filter-column metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelFilterColumnSnapshot {
        private readonly List<string> _values = new List<string>();

        /// <summary>
        /// Zero-based column identifier relative to the filter range.
        /// </summary>
        public int ColumnId { get; internal set; }

        /// <summary>
        /// Explicit visible values recorded by the source workbook for this filter.
        /// </summary>
        public IReadOnlyList<string> Values => _values;

        /// <summary>
        /// Custom filter predicates recorded for this column, when present.
        /// </summary>
        public ExcelCustomFiltersSnapshot? CustomFilters { get; internal set; }

        internal void AddValue(string value) {
            _values.Add(value ?? string.Empty);
        }
    }

    /// <summary>
    /// Immutable custom-filter metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelCustomFiltersSnapshot {
        private readonly List<ExcelCustomFilterConditionSnapshot> _conditions = new List<ExcelCustomFilterConditionSnapshot>();

        /// <summary>
        /// Whether all custom-filter conditions must match.
        /// </summary>
        public bool MatchAll { get; internal set; }

        /// <summary>
        /// Individual custom-filter conditions.
        /// </summary>
        public IReadOnlyList<ExcelCustomFilterConditionSnapshot> Conditions => _conditions;

        internal void AddCondition(ExcelCustomFilterConditionSnapshot condition) {
            if (condition == null) throw new ArgumentNullException(nameof(condition));
            _conditions.Add(condition);
        }
    }

    /// <summary>
    /// Immutable single custom-filter predicate discovered during inspection.
    /// </summary>
    public sealed class ExcelCustomFilterConditionSnapshot {
        /// <summary>
        /// Custom-filter operator name, when present.
        /// </summary>
        public string? Operator { get; internal set; }

        /// <summary>
        /// Comparison value stored in the workbook.
        /// </summary>
        public string Value { get; internal set; } = string.Empty;
    }

    /// <summary>
    /// Immutable table metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelTableSnapshot {
        private readonly List<ExcelTableColumnSnapshot> _columns = new List<ExcelTableColumnSnapshot>();

        /// <summary>
        /// Table name.
        /// </summary>
        public string Name { get; internal set; } = string.Empty;

        /// <summary>
        /// Table range in A1 notation.
        /// </summary>
        public string A1Range { get; internal set; } = string.Empty;

        /// <summary>
        /// One-based starting row of the table range.
        /// </summary>
        public int StartRow { get; internal set; }

        /// <summary>
        /// One-based ending row of the table range.
        /// </summary>
        public int EndRow { get; internal set; }

        /// <summary>
        /// One-based starting column of the table range.
        /// </summary>
        public int StartColumn { get; internal set; }

        /// <summary>
        /// One-based ending column of the table range.
        /// </summary>
        public int EndColumn { get; internal set; }

        /// <summary>
        /// Table style name, when present.
        /// </summary>
        public string? StyleName { get; internal set; }

        /// <summary>
        /// Whether the table has a header row.
        /// </summary>
        public bool HasHeaderRow { get; internal set; }

        /// <summary>
        /// Whether the table shows a totals/footer row.
        /// </summary>
        public bool TotalsRowShown { get; internal set; }

        /// <summary>
        /// Table-scoped auto filter metadata, when present.
        /// </summary>
        public ExcelAutoFilterSnapshot? AutoFilter { get; internal set; }

        /// <summary>
        /// Table columns discovered during inspection.
        /// </summary>
        public IReadOnlyList<ExcelTableColumnSnapshot> Columns => _columns;

        internal void AddColumn(ExcelTableColumnSnapshot column) {
            if (column == null) throw new ArgumentNullException(nameof(column));
            _columns.Add(column);
        }
    }

    /// <summary>
    /// Immutable table-column metadata discovered during inspection.
    /// </summary>
    public sealed class ExcelTableColumnSnapshot {
        /// <summary>
        /// One-based table column ordinal.
        /// </summary>
        public int Index { get; internal set; }

        /// <summary>
        /// Column name shown in the table header.
        /// </summary>
        public string Name { get; internal set; } = string.Empty;

        /// <summary>
        /// Totals-row function assigned to the column, when present.
        /// </summary>
        public string? TotalsRowFunction { get; internal set; }
    }

    /// <summary>
    /// Immutable named-range inspection snapshot.
    /// </summary>
    public sealed class ExcelNamedRangeSnapshot {
        /// <summary>
        /// Defined name.
        /// </summary>
        public string Name { get; internal set; } = string.Empty;

        /// <summary>
        /// Referenced A1 range or formula text stored for the name.
        /// </summary>
        public string ReferenceA1 { get; internal set; } = string.Empty;

        /// <summary>
        /// Sheet name for a local defined name, or <see langword="null"/> for workbook-global names.
        /// </summary>
        public string? SheetName { get; internal set; }

        /// <summary>
        /// Whether the name is an Excel built-in such as <c>_xlnm.Print_Area</c>.
        /// </summary>
        public bool IsBuiltIn { get; internal set; }
    }

    /// <summary>
    /// Immutable style metadata resolved for a cell.
    /// </summary>
    public sealed class ExcelCellStyleSnapshot {
        /// <summary>
        /// Original style index.
        /// </summary>
        public uint StyleIndex { get; internal set; }

        /// <summary>
        /// Number format identifier referenced by the style.
        /// </summary>
        public uint NumberFormatId { get; internal set; }

        /// <summary>
        /// Resolved number format code, when available.
        /// </summary>
        public string? NumberFormatCode { get; internal set; }

        /// <summary>
        /// Whether the number format looks date-like.
        /// </summary>
        public bool IsDateLike { get; internal set; }

        /// <summary>
        /// Whether the font is bold.
        /// </summary>
        public bool Bold { get; internal set; }

        /// <summary>
        /// Whether the font is italic.
        /// </summary>
        public bool Italic { get; internal set; }

        /// <summary>
        /// Whether the font is underlined.
        /// </summary>
        public bool Underline { get; internal set; }

        /// <summary>
        /// Font color in ARGB hexadecimal form, when directly resolvable.
        /// </summary>
        public string? FontColorArgb { get; internal set; }

        /// <summary>
        /// Fill color in ARGB hexadecimal form, when directly resolvable.
        /// </summary>
        public string? FillColorArgb { get; internal set; }

        /// <summary>
        /// Border metadata resolved for the cell style, when available.
        /// </summary>
        public ExcelCellBorderSnapshot? Border { get; internal set; }

        /// <summary>
        /// Horizontal alignment value, when specified.
        /// </summary>
        public string? HorizontalAlignment { get; internal set; }

        /// <summary>
        /// Vertical alignment value, when specified.
        /// </summary>
        public string? VerticalAlignment { get; internal set; }

        /// <summary>
        /// Whether wrap text is enabled.
        /// </summary>
        public bool WrapText { get; internal set; }
    }

    /// <summary>
    /// Immutable border metadata resolved for a cell style.
    /// </summary>
    public sealed class ExcelCellBorderSnapshot {
        /// <summary>
        /// Left border side.
        /// </summary>
        public ExcelBorderSideSnapshot? Left { get; internal set; }

        /// <summary>
        /// Right border side.
        /// </summary>
        public ExcelBorderSideSnapshot? Right { get; internal set; }

        /// <summary>
        /// Top border side.
        /// </summary>
        public ExcelBorderSideSnapshot? Top { get; internal set; }

        /// <summary>
        /// Bottom border side.
        /// </summary>
        public ExcelBorderSideSnapshot? Bottom { get; internal set; }
    }

    /// <summary>
    /// Immutable single-border-side metadata resolved from an OpenXML border definition.
    /// </summary>
    public sealed class ExcelBorderSideSnapshot {
        /// <summary>
        /// Border style name as stored in the workbook.
        /// </summary>
        public string? Style { get; internal set; }

        /// <summary>
        /// Border color in ARGB hexadecimal form, when directly resolvable.
        /// </summary>
        public string? ColorArgb { get; internal set; }
    }

    /// <summary>
    /// Immutable hyperlink metadata resolved for a cell.
    /// </summary>
    public sealed class ExcelHyperlinkSnapshot {
        /// <summary>
        /// Whether the hyperlink targets an external address.
        /// </summary>
        public bool IsExternal { get; internal set; }

        /// <summary>
        /// Target URI or workbook location.
        /// </summary>
        public string Target { get; internal set; } = string.Empty;
    }
}
