namespace OfficeIMO.Markdown;

/// <summary>
/// Typed table row containing header or body cells.
/// </summary>
public sealed class TableRow : MarkdownObject {
    private readonly IReadOnlyList<TableCell> _cells;

    /// <summary>Creates a typed table row.</summary>
    public TableRow(IEnumerable<TableCell>? cells, bool isHeader, int rowIndex) {
        _cells = cells == null
            ? Array.Empty<TableCell>()
            : new List<TableCell>(cells);
        IsHeader = isHeader;
        RowIndex = rowIndex;
    }

    /// <summary>Cells in this row, preserving document column order.</summary>
    public IReadOnlyList<TableCell> Cells => _cells;

    /// <summary>Whether this is the table header row.</summary>
    public bool IsHeader { get; }

    /// <summary>Zero-based body row index, or <c>-1</c> for the header row.</summary>
    public int RowIndex { get; }

    /// <summary>Gets a cell by zero-based column index.</summary>
    public TableCell? GetCell(int columnIndex) =>
        columnIndex >= 0 && columnIndex < _cells.Count ? _cells[columnIndex] : null;
}
