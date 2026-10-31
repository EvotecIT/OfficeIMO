namespace OfficeIMO.Excel;

public sealed partial class ExcelRange {
    /// <summary>Writes a rectangular matrix whose dimensions exactly match this range.</summary>
    public ExcelRange SetValues(object?[,] values) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        int rows = LastRow - FirstRow + 1, columns = LastColumn - FirstColumn + 1;
        if (values.GetLength(0) != rows || values.GetLength(1) != columns)
            throw new ArgumentException("Values array dimensions must match the range.", nameof(values));
        var cells = new (int Row, int Column, object Value)[checked(rows * columns)];
        for (int row = 0; row < rows; row++)
            for (int column = 0; column < columns; column++)
                cells[row * columns + column] = (FirstRow + row, FirstColumn + column, values[row, column]!);
        Sheet.CellValues(cells);
        return this;
    }

    /// <summary>Gets an editable cell using one-based offsets from this range's top-left cell.</summary>
    public ExcelCell CellAt(int rowOffset, int columnOffset) {
        if (rowOffset < 1 || rowOffset > LastRow - FirstRow + 1) throw new ArgumentOutOfRangeException(nameof(rowOffset));
        if (columnOffset < 1 || columnOffset > LastColumn - FirstColumn + 1) throw new ArgumentOutOfRangeException(nameof(columnOffset));
        return Sheet.CellAt(FirstRow + rowOffset - 1, FirstColumn + columnOffset - 1);
    }
}
