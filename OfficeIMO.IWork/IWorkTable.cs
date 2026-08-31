using System.Globalization;

namespace OfficeIMO.IWork;

/// <summary>One materialized, non-empty cell shared by Pages, Numbers, and Keynote table projections.</summary>
public sealed class IWorkTableCell {
    internal IWorkTableCell(int row, int column, IWorkCellKind kind, object? value,
        string? formula = null, string? error = null, IWorkCellKind? valueKind = null,
        bool formulaIsComplete = false) {
        Row = row;
        Column = column;
        Kind = kind;
        ValueKind = valueKind ?? kind;
        Value = value;
        Formula = formula;
        FormulaIsComplete = formulaIsComplete;
        Error = error;
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
    /// <summary>Gets the reconstructed formula, including its leading equals sign, or a visible marker when incomplete.</summary>
    public string? Formula { get; }
    /// <summary>Gets whether <see cref="Formula"/> is a complete editable expression.</summary>
    public bool FormulaIsComplete { get; }
    /// <summary>Gets a cell-level decode error without failing the surrounding table.</summary>
    public string? Error { get; }
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
        string? accessibilityDescription = null) {
        Name = name;
        RowCount = rowCount;
        ColumnCount = columnCount;
        HeaderRowCount = headerRowCount;
        HeaderColumnCount = headerColumnCount;
        FooterRowCount = footerRowCount;
        DefaultRowHeight = defaultRowHeight;
        DefaultColumnWidth = defaultColumnWidth;
        MergedRanges = Array.AsReadOnly((mergedRanges ?? Array.Empty<IWorkTableMergeRange>()).ToArray());
        Geometry = geometry;
        AccessibilityDescription = accessibilityDescription;
        _cells = new Dictionary<long, IWorkTableCell>();
        foreach (IWorkTableCell cell in cells) _cells[Key(cell.Row, cell.Column)] = cell;
        Cells = Array.AsReadOnly(_cells.Values.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToArray());
    }

    /// <summary>Gets the source table name.</summary>
    public string Name { get; }
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
    /// <summary>Gets the default column width in source points.</summary>
    public double? DefaultColumnWidth { get; }
    /// <summary>Gets merged ranges in source order.</summary>
    public IReadOnlyList<IWorkTableMergeRange> MergedRanges { get; }
    /// <summary>Gets the table drawable geometry when present.</summary>
    public IWorkGeometry? Geometry { get; }
    /// <summary>Gets the source table accessibility description.</summary>
    public string? AccessibilityDescription { get; }
    /// <summary>Gets materialized non-empty or diagnostic cells.</summary>
    public IReadOnlyList<IWorkTableCell> Cells { get; }

    /// <summary>Returns a materialized cell at a one-based position, or null when the source cell is empty.</summary>
    public IWorkTableCell? GetCell(int row, int column) {
        if (row < 1 || row > RowCount) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1 || column > ColumnCount) throw new ArgumentOutOfRangeException(nameof(column));
        return _cells.TryGetValue(Key(row, column), out IWorkTableCell? cell) ? cell : null;
    }

    internal bool HasPopulatedCoveredMergeCells() {
        if (MergedRanges.Count == 0) return false;
        if (Internal.IWorkMergeRangeValidator.HasOverlaps(MergedRanges, ColumnCount)) return true;
        foreach (IWorkTableCell cell in _cells.Values) {
            foreach (IWorkTableMergeRange merge in MergedRanges) {
                if (cell.Row < merge.FirstRow || cell.Row > merge.LastRow
                    || cell.Column < merge.FirstColumn || cell.Column > merge.LastColumn
                    || cell.Row == merge.FirstRow && cell.Column == merge.FirstColumn) continue;
                return true;
            }
        }
        return false;
    }

    private static long Key(int row, int column) => ((long)row << 32) | (uint)column;
}
