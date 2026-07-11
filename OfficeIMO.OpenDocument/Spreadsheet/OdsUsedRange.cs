namespace OfficeIMO.OpenDocument;

/// <summary>A zero-based rectangular range containing used spreadsheet cells.</summary>
public readonly struct OdsUsedRange : IEquatable<OdsUsedRange> {
    /// <summary>Creates a used range.</summary>
    public OdsUsedRange(long firstRow, long firstColumn, long lastRow, long lastColumn) {
        if (firstRow < 0 || firstColumn < 0 || lastRow < firstRow || lastColumn < firstColumn) throw new ArgumentOutOfRangeException(nameof(firstRow));
        FirstRow = firstRow; FirstColumn = firstColumn; LastRow = lastRow; LastColumn = lastColumn;
    }
    /// <summary>First row, zero-based.</summary>
    public long FirstRow { get; }
    /// <summary>First column, zero-based.</summary>
    public long FirstColumn { get; }
    /// <summary>Last row, zero-based and inclusive.</summary>
    public long LastRow { get; }
    /// <summary>Last column, zero-based and inclusive.</summary>
    public long LastColumn { get; }
    /// <summary>Logical row count.</summary>
    public long RowCount => checked(LastRow - FirstRow + 1);
    /// <summary>Logical column count.</summary>
    public long ColumnCount => checked(LastColumn - FirstColumn + 1);
    /// <inheritdoc />
    public bool Equals(OdsUsedRange other) => FirstRow == other.FirstRow && FirstColumn == other.FirstColumn && LastRow == other.LastRow && LastColumn == other.LastColumn;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OdsUsedRange other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => FirstRow.GetHashCode() ^ FirstColumn.GetHashCode() ^ LastRow.GetHashCode() ^ LastColumn.GetHashCode();
}
