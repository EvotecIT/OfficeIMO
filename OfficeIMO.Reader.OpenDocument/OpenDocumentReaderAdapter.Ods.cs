using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    private static IEnumerable<ReaderChunk> ReadSpreadsheet(OdsDocument document, string sourceName, ReaderOptions options, ReaderOpenDocumentOptions formatOptions,
        CancellationToken cancellationToken) {
        int blockIndex = 0;
        IEnumerable<OdsSheet> selected = document.Sheets;
        if (!string.IsNullOrWhiteSpace(formatOptions.SheetName)) {
            selected = selected.Where(sheet => string.Equals(sheet.Name, formatOptions.SheetName, StringComparison.Ordinal));
        }
        foreach (OdsSheet sheet in selected) {
            cancellationToken.ThrowIfCancellationRequested();
            OdsUsedRange? used = sheet.UsedRange;
            if (!used.HasValue) continue;
            OdsUsedRange range = used.Value;
            if (!string.IsNullOrWhiteSpace(formatOptions.A1Range)) {
                (int firstRow, int firstColumn, int lastRow, int lastColumn) = ParseA1Range(formatOptions.A1Range!);
                range = new OdsUsedRange(firstRow - 1L, firstColumn - 1L, lastRow - 1L, lastColumn - 1L);
            }
            int maxRows = options.MaxTableRows > 0 ? options.MaxTableRows : 200;
            const int maxColumns = 256;
            long sourceColumns = range.ColumnCount;
            int columnCount = (int)Math.Min(sourceColumns, maxColumns);
            long headerRow = range.FirstRow;
            string[] columns = Enumerable.Range(0, columnCount).Select(index => {
                string value = formatOptions.HeadersInFirstRow ? sheet.GetValue(headerRow, range.FirstColumn + index).ToString() : string.Empty;
                return value.Length == 0 ? "Column " + (index + 1).ToString(CultureInfo.InvariantCulture) : value;
            }).ToArray();
            long dataStart = checked(range.FirstRow + (formatOptions.HeadersInFirstRow ? 1 : 0));
            long availableDataRows = Math.Max(0, checked(range.LastRow - dataStart + 1));
            int emittedRows = (int)Math.Min(availableDataRows, maxRows);
            var rows = new List<IReadOnlyList<string>>(emittedRows);
            for (int rowOffset = 0; rowOffset < emittedRows; rowOffset++) {
                cancellationToken.ThrowIfCancellationRequested();
                long row = dataStart + rowOffset;
                rows.Add(Enumerable.Range(0, columnCount)
                    .Select(columnOffset => sheet.GetValue(row, range.FirstColumn + columnOffset).ToString()).ToArray());
            }
            string a1Range = ToA1(range.FirstRow, range.FirstColumn) + ":" + ToA1(range.LastRow, range.LastColumn);
            var location = new ReaderLocation {
                Path = sourceName, BlockIndex = blockIndex, SourceBlockIndex = blockIndex, SourceBlockKind = "sheet",
                Sheet = sheet.Name, A1Range = a1Range
            };
            var table = new ReaderTable {
                Title = sheet.Name, Kind = "ods-sheet", Columns = columns, Rows = rows,
                TotalRowCount = checked((int)Math.Min(availableDataRows, int.MaxValue)),
                Truncated = availableDataRows > emittedRows || sourceColumns > maxColumns,
                Location = location
            };
            var warnings = new List<string>();
            if (availableDataRows > emittedRows) warnings.Add("Sheet rows were truncated due to MaxTableRows.");
            if (sourceColumns > maxColumns) warnings.Add("Sheet columns were truncated to 256 columns for bounded extraction.");
            yield return new ReaderChunk {
                Id = BuildId(sourceName, "sheet", blockIndex), Kind = ReaderInputKind.OpenDocument,
                Location = location,
                Text = string.Join(Environment.NewLine, rows.Select(row => string.Join("\t", row))),
                Markdown = BuildTableMarkdown(columns, rows), Tables = new[] { table },
                Warnings = warnings.Count == 0 ? null : warnings
            };
            blockIndex++;
        }
    }

    private static string ToA1(long row, long column) {
        long value = checked(column + 1);
        var letters = new StringBuilder();
        while (value > 0) {
            value--;
            letters.Insert(0, (char)('A' + value % 26));
            value /= 26;
        }
        return letters + checked(row + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static (int FirstRow, int FirstColumn, int LastRow, int LastColumn) ParseA1Range(string value) {
        string[] parts = value.Split(':');
        if (parts.Length is < 1 or > 2) throw new FormatException("The ODS range must use A1 or A1:B2 notation.");
        (int row, int column) first = ParseA1Cell(parts[0]);
        (int row, int column) last = parts.Length == 1 ? first : ParseA1Cell(parts[1]);
        if (last.row < first.row || last.column < first.column) {
            throw new FormatException("The ODS range end must not precede its start.");
        }
        return (first.row, first.column, last.row, last.column);
    }

    private static (int Row, int Column) ParseA1Cell(string value) {
        string cell = value.Trim().Replace("$", string.Empty);
        int index = 0;
        int column = 0;
        while (index < cell.Length && char.IsLetter(cell[index])) {
            column = checked(column * 26 + char.ToUpperInvariant(cell[index]) - 'A' + 1);
            index++;
        }
        if (column == 0 || index == cell.Length || !int.TryParse(
                cell.Substring(index), NumberStyles.None, CultureInfo.InvariantCulture, out int row) || row < 1) {
            throw new FormatException("The ODS range contains an invalid A1 cell reference.");
        }
        return (row, column);
    }
}
