using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word;

public partial class WordTable {
    /// <summary>Inserts a row at a zero-based logical index, including rows inside content controls and custom XML.</summary>
    public WordTableRow InsertRow(int index, params object?[] values) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        var existing = WordTableContent.Rows(_table).ToList();
        if (index < 0 || index > existing.Count) throw new ArgumentOutOfRangeException(nameof(index));
        int columns = existing.Count > 0 ? WordTableContent.Cells(existing[0]).Count() : values.Length;
        if (values.Length > columns) throw new ArgumentException("The inserted row has more values than the table has columns.", nameof(values));
        var row = AddRow(columns);
        if (index < existing.Count) {
            row._tableRow.Remove();
            WordTableContent.InsertRowBefore(_table, existing[index], row._tableRow);
            InvalidateRowElements();
        }
        for (int column = 0; column < values.Length; column++)
            row.Cells[column].AddParagraph(System.Convert.ToString(values[column], CultureInfo.CurrentCulture) ?? string.Empty, true);
        return row;
    }

    /// <summary>Inserts a column at a zero-based logical index and returns its cells in row order.</summary>
    public IReadOnlyList<WordTableCell> InsertColumn(int index, params object?[] values) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        var rows = Rows;
        var cells = rows.Select(row => row.Cells).ToArray();
        if (index < 0 || cells.Any(row => index > row.Count)) throw new ArgumentOutOfRangeException(nameof(index));
        if (values.Length > rows.Count) throw new ArgumentException("The inserted column has more values than the table has rows.", nameof(values));
        var added = new List<WordTableCell>(rows.Count);
        for (int row = 0; row < rows.Count; row++) {
            var cell = new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2400" }), new Paragraph());
            if (index < cells[row].Count) cells[row][index]._tableCell.InsertBeforeSelf(cell);
            else rows[row]._tableRow.Append(cell);
            var wrapper = new WordTableCell(_document, this, rows[row], cell);
            if (row < values.Length) wrapper.AddParagraph(System.Convert.ToString(values[row], CultureInfo.CurrentCulture) ?? string.Empty, true);
            added.Add(wrapper);
        }
        RefreshGrid();
        return added;
    }
}
