using System.Globalization;
using OfficeIMO.Reader;
using OfficeIMO.Word;

namespace OfficeIMO.Reader.Word;

internal static class WordTableProjection {
    internal static ReaderTable Map(
        WordTableSnapshot table,
        ReaderLocation location,
        int tableIndex,
        int maxRows) {
        int columnCount = Math.Max(
            table.ColumnCount,
            table.Rows.Count == 0 ? 0 : table.Rows.Max(static row => row.Cells.Count));
        bool hasHeaderRow = table.RepeatHeaderRow && table.Rows.Count > 0;
        IReadOnlyList<string> columns = hasHeaderRow
            ? BuildRowValues(table.Rows[0], columnCount, useFallbacks: true)
            : BuildFallbackColumns(columnCount);
        int dataStart = hasHeaderRow ? 1 : 0;
        int totalRowCount = Math.Max(0, table.Rows.Count - dataStart);
        IEnumerable<WordTableRowSnapshot> sourceRows = table.Rows.Skip(dataStart);
        bool truncated = maxRows > 0 && totalRowCount > maxRows;
        if (truncated) sourceRows = sourceRows.Take(maxRows);
        IReadOnlyList<IReadOnlyList<string>> rows = sourceRows
            .Select(row => BuildRowValues(row, columnCount, useFallbacks: false))
            .ToArray();
        string? title = !string.IsNullOrWhiteSpace(table.Title)
            ? table.Title
            : !string.IsNullOrWhiteSpace(table.Description)
                ? table.Description
                : "Word table " + (tableIndex + 1).ToString(CultureInfo.InvariantCulture);
        return new ReaderTable {
            Title = title,
            Kind = "word-table",
            Location = location,
            Columns = columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows),
            Rows = rows,
            TotalRowCount = totalRowCount,
            Truncated = truncated
        };
    }

    private static IReadOnlyList<string> BuildRowValues(
        WordTableRowSnapshot row,
        int columnCount,
        bool useFallbacks) {
        var values = new string[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            WordTableCellSnapshot? cell = row.Cells.FirstOrDefault(candidate => candidate.ColumnIndex == columnIndex)
                ?? (columnIndex < row.Cells.Count ? row.Cells[columnIndex] : null);
            string value = cell == null
                ? string.Empty
                : string.Join(" ", cell.Paragraphs
                    .Select(static paragraph => paragraph.Text)
                    .Where(static text => !string.IsNullOrWhiteSpace(text)));
            values[columnIndex] = string.IsNullOrWhiteSpace(value) && useFallbacks
                ? "Column " + (columnIndex + 1).ToString(CultureInfo.InvariantCulture)
                : value;
        }
        return values;
    }

    private static IReadOnlyList<string> BuildFallbackColumns(int count) =>
        Enumerable.Range(1, count)
            .Select(index => "Column " + index.ToString(CultureInfo.InvariantCulture))
            .ToArray();
}
