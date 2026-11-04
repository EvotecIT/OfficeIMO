using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word;

public partial class WordTable {
    /// <summary>
    /// Inserts a row at a zero-based logical index, including rows inside content controls and custom XML.
    /// The row spans the table's grid columns; inside a vertical merge it joins the merged cell, and that column's
    /// value is added to the merged cell.
    /// </summary>
    public WordTableRow InsertRow(int index, params object?[] values) {
        ArgumentNullException.ThrowIfNull(values);
        var existing = WordTableContent.Rows(_table).ToList();
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(index, existing.Count);
        int columns = existing.Count > 0 ? WordTableContent.GridColumnCount(existing[0]) : values.Length;
        if (values.Length > columns) throw new ArgumentException("The inserted row has more values than the table has columns.", nameof(values));
        var row = AddRow(columns);
        if (index < existing.Count) {
            row._tableRow.Remove();
            WordTableContent.InsertRowBefore(_table, existing[index], row._tableRow);
            InvalidateRowElements();
            ContinueVerticalMerges(row, existing[index]);
        }
        for (int column = 0; column < values.Length; column++) {
            string text = System.Convert.ToString(values[column], CultureInfo.CurrentCulture) ?? string.Empty;
            TableCell? merged = index < existing.Count ? FindVerticalMergeStart(existing, index, column) : null;
            if (merged != null) GetCell(merged).AddParagraph(text);
            else GetCell(WordTableContent.CellAtColumn(row._tableRow, column)!.Value.Cell).AddParagraph(text, true);
        }
        return row;
    }

    /// <summary>
    /// Inserts a column at a zero-based grid column and returns its cells in row order. Inside a merged or spanned
    /// cell the column widens that cell, which is the cell returned for the row and receives its value.
    /// </summary>
    public IReadOnlyList<WordTableCell> InsertColumn(int index, params object?[] values) {
        ArgumentNullException.ThrowIfNull(values);
        var rows = Rows;
        if (index < 0 || rows.Any(row => index > WordTableContent.GridColumnCount(row._tableRow))) throw new ArgumentOutOfRangeException(nameof(index));
        if (values.Length > rows.Count) throw new ArgumentException("The inserted column has more values than the table has rows.", nameof(values));
        var added = new List<WordTableCell>(rows.Count);
        for (int row = 0; row < rows.Count; row++) {
            WordTableCell cell = InsertColumnCell(rows[row], index, out bool merged);
            if (row < values.Length) cell.AddParagraph(System.Convert.ToString(values[row], CultureInfo.CurrentCulture) ?? string.Empty, !merged);
            added.Add(cell);
        }
        RefreshGrid();
        return added;
    }

    private WordTableCell InsertColumnCell(WordTableRow row, int column, out bool merged) {
        var cell = new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2400" }), new Paragraph());
        merged = false;
        if (WordTableContent.CellAtColumn(row._tableRow, column) is not (TableCell target, int start)) {
            row._tableRow.Append(cell);
        } else if (start < column) {
            target.TableCellProperties!.GridSpan = new GridSpan { Val = WordTableContent.GridSpanOf(target) + 1 };
            merged = true;
            return new WordTableCell(_document, this, row, target);
        } else {
            target.InsertBeforeSelf(cell);
            if (WordTableContent.ContinuesHorizontalMerge(target)) {
                cell.TableCellProperties!.HorizontalMerge = new HorizontalMerge { Val = MergedCellValues.Continue };
                merged = true;
                TableCell anchor = target;
                while (WordTableContent.ContinuesHorizontalMerge(anchor) && anchor.PreviousSibling<TableCell>() is { } previous) anchor = previous;
                return new WordTableCell(_document, this, row, anchor);
            }
        }
        return new WordTableCell(_document, this, row, cell);
    }

    private WordTableCell GetCell(TableCell cell) {
        TableRow tableRow = cell.Ancestors<TableRow>().First();
        return Rows.First(row => row._tableRow == tableRow).Cells.First(candidate => candidate._tableCell == cell);
    }

    private static void ContinueVerticalMerges(WordTableRow row, TableRow next) {
        TableCell[] cells = row.Cells.Select(cell => cell._tableCell).ToArray();
        for (int column = 0; column < cells.Length;) {
            if (WordTableContent.CellAtColumn(next, column) is not (TableCell below, int start)) break;
            int span = WordTableContent.GridSpanOf(below), covered = start + span - column;
            if (WordTableContent.ContinuesVerticalMerge(below)) {
                cells[column].TableCellProperties!.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Continue };
                if (below.TableCellProperties?.HorizontalMerge is { } horizontal) {
                    cells[column].TableCellProperties!.HorizontalMerge = (HorizontalMerge)horizontal.CloneNode(true);
                }
                if (span > 1) {
                    cells[column].TableCellProperties!.GridSpan = new GridSpan { Val = span };
                    for (int extra = column + 1; extra < column + covered && extra < cells.Length; extra++) cells[extra].Remove();
                }
            }
            column += covered;
        }
    }

    private static TableCell? FindVerticalMergeStart(IReadOnlyList<TableRow> rows, int index, int column) {
        if (!WordTableContent.ContinuesVerticalMerge(WordTableContent.CellAtColumn(rows[index], column)?.Cell)) return null;
        for (int row = index - 1; row >= 0; row--) {
            TableCell? cell = WordTableContent.CellAtColumn(rows[row], column)?.Cell;
            if (!WordTableContent.ContinuesVerticalMerge(cell)) return cell;
        }
        return null;
    }

    /// <summary>Hands a merged cell starting in a row about to be removed, with its content, to the next row.</summary>
    internal void ReleaseVerticalMerges(TableRow row) {
        var rows = WordTableContent.Rows(_table).ToList();
        int index = rows.IndexOf(row);
        if (index < 0 || index + 1 >= rows.Count) return;
        int column = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        foreach (TableCell cell in WordTableContent.Cells(row).ToList()) {
            int span = WordTableContent.GridSpanOf(cell);
            TableCell? below = WordTableContent.CellAtColumn(rows[index + 1], column)?.Cell;
            if (cell.TableCellProperties?.VerticalMerge != null && !WordTableContent.ContinuesVerticalMerge(cell)
                && WordTableContent.ContinuesVerticalMerge(below)) {
                foreach (var content in below!.ChildElements.Where(child => child is not TableCellProperties).ToList()) content.Remove();
                foreach (var content in cell.ChildElements.Where(child => child is not TableCellProperties).ToList()) {
                    content.Remove();
                    below.Append(content);
                }
                below.TableCellProperties!.VerticalMerge!.Val = MergedCellValues.Restart;
            }
            column += span;
        }
    }
}
