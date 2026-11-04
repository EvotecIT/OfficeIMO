using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>An immutable border shared by cells along a horizontal or vertical range of table grid boundaries.</summary>
public readonly struct PublisherTableBorderSegment : IEquatable<PublisherTableBorderSegment> {
    internal PublisherTableBorderSegment(int firstRow, int firstColumn, int lastRow, int lastColumn, OfficeBorderSide side) {
        FirstRow = firstRow; FirstColumn = firstColumn; LastRow = lastRow; LastColumn = lastColumn; Side = side;
    }
    /// <summary>Starting row boundary, from zero through the table's row count.</summary>
    public int FirstRow { get; }
    /// <summary>Starting column boundary, from zero through the table's column count.</summary>
    public int FirstColumn { get; }
    /// <summary>Ending row boundary. Horizontal segments keep the starting row.</summary>
    public int LastRow { get; }
    /// <summary>Ending column boundary. Vertical segments keep the starting column.</summary>
    public int LastColumn { get; }
    /// <summary>Border color and width in points.</summary>
    public OfficeBorderSide Side { get; }
    /// <inheritdoc />
    public bool Equals(PublisherTableBorderSegment other) => FirstRow == other.FirstRow && FirstColumn == other.FirstColumn
        && LastRow == other.LastRow && LastColumn == other.LastColumn && Side.Equals(other.Side);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is PublisherTableBorderSegment other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked(((((FirstRow * 397) ^ FirstColumn) * 397 ^ LastRow) * 397 ^ LastColumn) * 397 ^ Side.GetHashCode());
}
