using System.Globalization;

namespace OfficeIMO.IWork;

/// <summary>One materialized value or explicitly styled empty cell shared by Pages, Numbers, and Keynote table projections.</summary>
public sealed class IWorkTableCell {
    internal IWorkTableCell(int row, int column, IWorkCellKind kind, object? value,
        string? formula = null, string? error = null, IWorkCellKind? valueKind = null,
        bool formulaIsComplete = false, IWorkTextContent? richText = null,
        bool cachedValueIsComplete = true, bool sourceFormulaIsDeclared = false,
        bool hasDecodeError = false, IWorkNumberFormat? numberFormat = null,
        string? sourceNumberText = null, bool numericValueIsApproximate = false, Internal.IWorkFormulaDefinition? formulaDefinition = null, IWorkCellFill? fill = null,
        IWorkCellPadding? padding = null, IWorkCellVerticalAlignment? verticalAlignment = null) {
        Row = row;
        Column = column;
        Kind = kind;
        ValueKind = valueKind ?? kind;
        Value = value;
        Formula = formula;
        FormulaIsComplete = formulaIsComplete;
        SourceFormulaIsDeclared = sourceFormulaIsDeclared || kind == IWorkCellKind.Formula;
        CachedValueIsComplete = cachedValueIsComplete;
        Error = error;
        HasDecodeError = hasDecodeError;
        RichText = richText;
        NumberFormat = numberFormat;
        SourceNumberText = sourceNumberText;
        NumericValueIsApproximate = numericValueIsApproximate;
        FormulaDefinition = formulaDefinition;
        Fill = fill;
        Padding = padding;
        VerticalAlignment = verticalAlignment;
    }

    /// <summary>Gets the one-based row position.</summary>
    public int Row { get; }
    /// <summary>Gets the one-based column position.</summary>
    public int Column { get; }
    /// <summary>Gets the recovered value kind.</summary>
    public IWorkCellKind Kind { get; }
    /// <summary>Gets the type of <see cref="Value"/>, including the cached-value type of a formula cell.</summary>
    public IWorkCellKind ValueKind { get; }
    /// <summary>Gets the typed cached value, when one was recovered.</summary>
    public object? Value { get; }
    /// <summary>Gets the reconstructed source formula, including its leading equals sign, or a visible marker when incomplete.</summary>
    /// <remarks>Bound Numbers table references use quoted source sheet and table labels.
    /// Destination adapters render those references using their actual destination names.</remarks>
    public string? Formula { get; }
    /// <summary>Gets whether <see cref="Formula"/> is a complete reconstructed source expression.</summary>
    public bool FormulaIsComplete { get; }
    /// <summary>Gets whether a supported source cell header declares a formula, even when its contents could not be decoded.</summary>
    public bool SourceFormulaIsDeclared { get; }
    /// <summary>Gets whether the recovered cached value is complete.</summary>
    public bool CachedValueIsComplete { get; }
    /// <summary>Gets an error marker or cell-level decode failure without failing the surrounding table.</summary>
    public string? Error { get; }
    /// <summary>Gets whether a storage or value decoding failure replaced this cell with an error. A recovered native error marker does not set this flag; false does not establish complete field fidelity.</summary>
    public bool HasDecodeError { get; }
    /// <summary>Gets source rich text for a text or formula cell, including runs, styles, and hyperlinks when recovered.</summary>
    public IWorkTextContent? RichText { get; }
    /// <summary>Gets the supported source numeric format. Null means absent or unresolved; source diagnostics distinguish unsupported declarations. Raw display text does not apply this format.</summary>
    public IWorkNumberFormat? NumberFormat { get; }
    /// <summary>Gets the exact normalized source Decimal128 value in invariant coefficient/exponent notation when <see cref="NumericValueIsApproximate"/> is true. This includes decimal formula caches.</summary>
    public string? SourceNumberText { get; }
    /// <summary>Gets whether a source Decimal128 value exceeds the fifteen-significant-digit portable numeric contract. The recovered double remains available and conversion reports an approximation.</summary>
    public bool NumericValueIsApproximate { get; }

    /// <summary>Gets the supported selected native cell fill. Null means absent or unresolved;
    /// <see cref="IWorkCellFill.IsNone"/> distinguishes an explicit no-fill override.</summary>
    public IWorkCellFill? Fill { get; }

    /// <summary>Gets selected native padding in points. Null means absent or unresolved; an explicit empty message retains zero on all four sides.</summary>
    public IWorkCellPadding? Padding { get; }

    /// <summary>Gets selected native vertical alignment. Null means absent or unresolved.</summary>
    public IWorkCellVerticalAlignment? VerticalAlignment { get; }

    internal bool HasCellFormatting => Fill != null || Padding != null || VerticalAlignment != null;

    internal Internal.IWorkFormulaDefinition? FormulaDefinition { get; }

    internal IWorkTableCell WithFormula(Internal.IWorkFormulaResult result) =>
        new(Row, Column, Kind, Value, result.Text.Length == 0 ? "=?" : result.Text, Error, ValueKind, result.IsComplete,
            RichText, CachedValueIsComplete, SourceFormulaIsDeclared, HasDecodeError, NumberFormat,
            SourceNumberText, NumericValueIsApproximate, FormulaDefinition, Fill, Padding, VerticalAlignment);

    internal IWorkTableCell WithNumberFormat(IWorkNumberFormat format) =>
        new(Row, Column, Kind, Value, Formula, Error, ValueKind, FormulaIsComplete,
            RichText, CachedValueIsComplete, SourceFormulaIsDeclared, HasDecodeError, format,
            SourceNumberText, NumericValueIsApproximate, FormulaDefinition, Fill, Padding, VerticalAlignment);

    internal IWorkTableCell WithSourceNumber(string text, bool approximate) =>
        new(Row, Column, Kind, Value, Formula, Error, ValueKind, FormulaIsComplete,
            RichText, CachedValueIsComplete, SourceFormulaIsDeclared, HasDecodeError, NumberFormat,
            text, approximate, FormulaDefinition, Fill, Padding, VerticalAlignment);
    internal IWorkTableCell WithStyle(Internal.IWorkTableCellStyle style) =>
        new(Row, Column, Kind, Value, Formula, Error, ValueKind, FormulaIsComplete,
            RichText, CachedValueIsComplete, SourceFormulaIsDeclared, HasDecodeError, NumberFormat,
            SourceNumberText, NumericValueIsApproximate, FormulaDefinition, style.Fill, style.Padding, style.VerticalAlignment);

    /// <summary>Gets a culture-invariant display representation of the recovered value or formula.</summary>
    public string DisplayText => Kind switch {
        IWorkCellKind.Boolean => Convert.ToBoolean(Value, CultureInfo.InvariantCulture) ? "TRUE" : "FALSE",
        IWorkCellKind.DateTime when Value is DateTime date => FormatDateTime(date),
        IWorkCellKind.Duration when Value is double seconds => seconds.ToString("R", CultureInfo.InvariantCulture) + "s",
        IWorkCellKind.Formula => Formula ?? "=?",
        IWorkCellKind.Error => Error ?? "#ERROR",
        _ => Convert.ToString(Value, CultureInfo.InvariantCulture) ?? string.Empty
    };

    /// <summary>Gets a culture-invariant display representation of the cached value, including for formula cells.</summary>
    public string CachedDisplayText => ValueKind switch {
        IWorkCellKind.Boolean => Convert.ToBoolean(Value, CultureInfo.InvariantCulture) ? "TRUE" : "FALSE",
        IWorkCellKind.DateTime when Value is DateTime date => FormatDateTime(date),
        IWorkCellKind.Duration when Value is double seconds => seconds.ToString("R", CultureInfo.InvariantCulture) + "s",
        IWorkCellKind.Error => Error ?? "#ERROR",
        _ => Convert.ToString(Value, CultureInfo.InvariantCulture) ?? string.Empty
    };

    private static string FormatDateTime(DateTime value) =>
        value.ToString("yyyy-MM-dd HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture);
}

/// <summary>One rectangular merged-cell range in an iWork table.</summary>
public sealed class IWorkTableMergeRange {
    internal IWorkTableMergeRange(int firstRow, int firstColumn, int lastRow, int lastColumn) {
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        LastRow = lastRow;
        LastColumn = lastColumn;
    }

    /// <summary>Gets the one-based first row.</summary>
    public int FirstRow { get; }
    /// <summary>Gets the one-based first column.</summary>
    public int FirstColumn { get; }
    /// <summary>Gets the one-based last row.</summary>
    public int LastRow { get; }
    /// <summary>Gets the one-based last column.</summary>
    public int LastColumn { get; }
}

/// <summary>A sparse table shared by Pages, Numbers, and Keynote projections.</summary>
public sealed class IWorkTable {
    private readonly Dictionary<long, IWorkTableCell> _cells;

    internal IWorkTable(string name, int rowCount, int columnCount,
        IReadOnlyList<IWorkTableCell> cells, int headerRowCount = 0, int headerColumnCount = 0,
        int footerRowCount = 0, double? defaultRowHeight = null, double? defaultColumnWidth = null,
        IReadOnlyList<IWorkTableMergeRange>? mergedRanges = null, IWorkGeometry? geometry = null,
        string? accessibilityDescription = null, IWorkObjectIdentity? sourceIdentity = null,
        IReadOnlyList<IWorkObjectIdentity>? omittedTextUnits = null,
        IReadOnlyDictionary<int, double>? rowHeights = null,
        IReadOnlyDictionary<int, double>? columnWidths = null,
        Guid? formulaIdentifier = null, IWorkArchiveRecord? modelRecord = null, bool bodyMetadataIsComplete = true, bool? autoResizeRows = null) {
        Name = name;
        FormulaIdentifier = formulaIdentifier;
        ModelRecord = modelRecord;
        BodyMetadataIsComplete = bodyMetadataIsComplete;
        RowCount = rowCount;
        ColumnCount = columnCount;
        HeaderRowCount = headerRowCount;
        HeaderColumnCount = headerColumnCount;
        FooterRowCount = footerRowCount;
        DefaultRowHeight = defaultRowHeight;
        AutoResizeRows = autoResizeRows;
        DefaultColumnWidth = defaultColumnWidth;
        RowHeights = CopyDimensions(rowHeights);
        ColumnWidths = CopyDimensions(columnWidths);
        MergedRanges = Array.AsReadOnly((mergedRanges ?? Array.Empty<IWorkTableMergeRange>()).ToArray());
        Geometry = geometry;
        AccessibilityDescription = accessibilityDescription;
        SourceIdentity = sourceIdentity;
        OmittedTextUnits = Array.AsReadOnly((omittedTextUnits ?? Array.Empty<IWorkObjectIdentity>()).ToArray());
        _cells = new Dictionary<long, IWorkTableCell>();
        foreach (IWorkTableCell cell in cells) _cells[Key(cell.Row, cell.Column)] = cell;
        Cells = Array.AsReadOnly(_cells.Values.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToArray());
    }

    internal bool BodyMetadataIsComplete { get; }
    internal Guid? FormulaIdentifier { get; }
    internal IWorkArchiveRecord? ModelRecord { get; }
    internal IWorkTable WithCells(IReadOnlyList<IWorkTableCell> cells) =>
        new(Name, RowCount, ColumnCount, cells, HeaderRowCount, HeaderColumnCount, FooterRowCount,
            DefaultRowHeight, DefaultColumnWidth, MergedRanges, Geometry, AccessibilityDescription,
            SourceIdentity, OmittedTextUnits, RowHeights, ColumnWidths, FormulaIdentifier, ModelRecord, BodyMetadataIsComplete, AutoResizeRows);

    /// <summary>Gets the source table name.</summary>
    public string Name { get; }
    /// <summary>Gets the native table-info identity, excluding auxiliary model and tile records.</summary>
    public IWorkObjectIdentity? SourceIdentity { get; }
    internal IReadOnlyList<IWorkObjectIdentity> OmittedTextUnits { get; }
    /// <summary>Gets the declared row count without allocating an equivalent dense grid.</summary>
    public int RowCount { get; }
    /// <summary>Gets the declared column count without allocating an equivalent dense grid.</summary>
    public int ColumnCount { get; }
    /// <summary>Gets the declared leading header-row count.</summary>
    public int HeaderRowCount { get; }
    /// <summary>Gets the declared leading header-column count.</summary>
    public int HeaderColumnCount { get; }
    /// <summary>Gets the declared trailing footer-row count.</summary>
    public int FooterRowCount { get; }
    /// <summary>Gets the default row height in source points.</summary>
    public double? DefaultRowHeight { get; }
    /// <summary>Gets whether native table rows grow to fit content. Null means the setting is absent or unresolved.</summary>
    public bool? AutoResizeRows { get; }
    /// <summary>Gets the default column width in source points.</summary>
    public double? DefaultColumnWidth { get; }
    /// <summary>Gets explicit row heights in points, keyed by one-based row position. Zero-size native entries use the default and are omitted.</summary>
    public IReadOnlyDictionary<int, double> RowHeights { get; }
    /// <summary>Gets explicit column widths in points, keyed by one-based column position. Zero-size native entries use the default and are omitted.</summary>
    public IReadOnlyDictionary<int, double> ColumnWidths { get; }

    /// <summary>Gets the explicit height or default height of a one-based row, when known.</summary>
    public double? GetRowHeight(int row) {
        if (row < 1 || row > RowCount) throw new ArgumentOutOfRangeException(nameof(row));
        return RowHeights.TryGetValue(row, out double height) ? height : DefaultRowHeight;
    }

    /// <summary>Gets the explicit width or default width of a one-based column, when known.</summary>
    public double? GetColumnWidth(int column) {
        if (column < 1 || column > ColumnCount) throw new ArgumentOutOfRangeException(nameof(column));
        return ColumnWidths.TryGetValue(column, out double width) ? width : DefaultColumnWidth;
    }

    /// <summary>Gets merged ranges in source order.</summary>
    public IReadOnlyList<IWorkTableMergeRange> MergedRanges { get; }
    /// <summary>Gets the table drawable geometry when present.</summary>
    public IWorkGeometry? Geometry { get; }
    /// <summary>Gets the source table accessibility description.</summary>
    public string? AccessibilityDescription { get; }
    /// <summary>Gets materialized value, diagnostic, or explicitly styled empty cells.</summary>
    public IReadOnlyList<IWorkTableCell> Cells { get; }

    /// <summary>Returns a materialized cell at a one-based position, or null when no value, diagnostic, or supported cell formatting is retained.</summary>
    public IWorkTableCell? GetCell(int row, int column) {
        if (row < 1 || row > RowCount) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1 || column > ColumnCount) throw new ArgumentOutOfRangeException(nameof(column));
        return _cells.TryGetValue(Key(row, column), out IWorkTableCell? cell) ? cell : null;
    }

    internal bool HasPopulatedCoveredMergeCells() {
        return Internal.IWorkMergeRangeValidator.HasOverlapsOrCoveredCells(
            MergedRanges, Cells, ColumnCount);
    }

    private static IReadOnlyDictionary<int, double> CopyDimensions(IReadOnlyDictionary<int, double>? source) {
        var copy = new Dictionary<int, double>();
        if (source != null) foreach (var pair in source) copy.Add(pair.Key, pair.Value);
        return new System.Collections.ObjectModel.ReadOnlyDictionary<int, double>(copy);
    }

    private static long Key(int row, int column) => ((long)row << 32) | (uint)column;
}
