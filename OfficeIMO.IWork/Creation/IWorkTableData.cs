namespace OfficeIMO.IWork;

/// <summary>Mutable sparse scalar data for a newly created iWork table.</summary>
/// <remarks>Coordinates are one-based. Each axis supports 1–256 positions. Unset positions are empty;
/// text is stored as text even when it begins with an equals sign. Instances are not thread safe.</remarks>
public sealed class IWorkTableData {
    private readonly Dictionary<int, IWorkTableCell> _cells = new();

    private IWorkTableData(int rows, int columns) {
        if (rows is < 1 or > 256) throw new ArgumentOutOfRangeException(nameof(rows), "Between 1 and 256 rows are supported.");
        if (columns is < 1 or > 256) throw new ArgumentOutOfRangeException(nameof(columns), "Between 1 and 256 columns are supported.");
        RowCount = rows; ColumnCount = columns;
    }

    /// <summary>Gets the number of rows, including unset empty positions.</summary>
    public int RowCount { get; }
    /// <summary>Gets the number of columns, including unset empty positions.</summary>
    public int ColumnCount { get; }
    /// <summary>Creates an empty table with fixed row and column counts.</summary>
    public static IWorkTableData Create(int rows, int columns) => new(rows, columns);

    /// <summary>Sets text, normalizing CR/CRLF to LF and rejecting other controls and invalid UTF-16.</summary>
    public IWorkTableData SetText(int row, int column, string text) =>
        Set(row, column, IWorkCellKind.Text, IWorkCreationValidation.NormalizeText(text));

    /// <summary>Sets a finite IEEE-754 double as a numeric value.</summary>
    public IWorkTableData SetNumber(int row, int column, double value) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value), "A finite number is required.");
        return Set(row, column, IWorkCellKind.Number, value);
    }

    /// <summary>Sets a boolean value, retaining its type rather than a text label.</summary>
    public IWorkTableData SetBoolean(int row, int column, bool value) => Set(row, column, IWorkCellKind.Boolean, value);

    /// <summary>Returns a position to an unset empty cell.</summary>
    public IWorkTableData ClearCell(int row, int column) {
        _cells.Remove(Key(row, column)); return this;
    }

    /// <summary>Gets the scalar at a position, or an empty cell when it has not been set.</summary>
    public IWorkTableCell GetCell(int row, int column) => _cells.TryGetValue(Key(row, column), out var cell)
        ? cell : new IWorkTableCell(row, column, IWorkCellKind.Empty, null);

    internal IWorkTableCell[] Snapshot() => _cells.OrderBy(entry => entry.Key).Select(entry => entry.Value).ToArray();

    private IWorkTableData Set(int row, int column, IWorkCellKind kind, object value) {
        _cells[Key(row, column)] = new IWorkTableCell(row, column, kind, value); return this;
    }

    private int Key(int row, int column) {
        if (row < 1 || row > RowCount) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1 || column > ColumnCount) throw new ArgumentOutOfRangeException(nameof(column));
        return (row - 1) * ColumnCount + column - 1;
    }
}
