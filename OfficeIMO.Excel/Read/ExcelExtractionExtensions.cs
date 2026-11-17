using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel;

/// <summary>
/// Chunked extraction helpers intended for AI ingestion.
/// </summary>
internal static class ExcelExtractionExtensions {
    /// <summary>
    /// Options controlling Excel extraction behavior.
    /// </summary>
    public sealed class ExcelExtractOptions {
        /// <summary>
        /// When true, the first row of the range is treated as headers. Default: true.
        /// </summary>
        public bool HeadersInFirstRow { get; set; } = true;

        /// <summary>
        /// Number of worksheet rows per emitted chunk when streaming. Default: 200.
        /// </summary>
        public int ChunkRows { get; set; } = 200;

        /// <summary>
        /// When true, emit a Markdown preview table in <see cref="ExcelExtractChunk.Markdown"/>. Default: true.
        /// </summary>
        public bool EmitMarkdownTable { get; set; } = true;
    }

    /// <summary>
    /// Extracts an Excel sheet range into row-chunked <see cref="ExcelExtractChunk"/> instances.
    /// </summary>
    /// <param name="reader">Workbook reader.</param>
    /// <param name="sheetName">Sheet to extract.</param>
    /// <param name="a1Range">A1 range; when null, uses the sheet's used range.</param>
    /// <param name="extract">Extraction options.</param>
    /// <param name="chunking">Chunking options.</param>
    /// <param name="sourcePath">Optional source path for citations.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ExcelExtractChunk> ExtractChunks(
        this ExcelDocumentReader reader,
        string sheetName,
        string? a1Range = null,
        ExcelExtractOptions? extract = null,
        ExcelExtractChunkingOptions? chunking = null,
        string? sourcePath = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(reader);
        if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentNullException(nameof(sheetName));

        extract ??= new ExcelExtractOptions();
        chunking ??= new ExcelExtractChunkingOptions();
        if (extract.ChunkRows < 1) extract.ChunkRows = 1;
        if (chunking.MaxChars < 256) chunking.MaxChars = 256;
        if (chunking.MaxTableRows < 1) chunking.MaxTableRows = 1;

        var sheet = reader.GetSheet(sheetName);
        var range = string.IsNullOrWhiteSpace(a1Range) ? sheet.GetUsedRangeA1() : a1Range!;

        // Determine headers from the first row in the overall range.
        (int r1, int c1, int r2, int c2) = ExcelA1.ParseRange(range);
        var width = Math.Max(0, c2 - c1 + 1);
        long projectedCells = (long)Math.Max(0, r2 - r1 + 1) * width;
        if (projectedCells > chunking.MaxProjectedCells - chunking.ProjectedCells)
            throw new OfficeInspectionResourceLimitException("excel", nameof(chunking.MaxProjectedCells), chunking.MaxProjectedCells);
        chunking.ProjectedCells += projectedCells;
        var headers = new List<string>(width);
        bool headersResolved = false;

        int chunkIndex = 0;
        foreach (var chunk in sheet.ReadRangeStream(range, extract.ChunkRows, ct: cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();

            if (chunk.RowCount <= 0 || chunk.ColCount <= 0) {
                chunkIndex++;
                continue;
            }

            if (!headersResolved) {
                headersResolved = true;
                if (extract.HeadersInFirstRow && chunk.Rows.Length > 0) {
                    headers.Clear();
                    var hdrRow = chunk.Rows[0];
                    for (int c = 0; c < chunk.ColCount; c++) headers.Add(ToCellString(hdrRow[c], fallback: $"Column{c + 1}"));
                } else {
                    headers.Clear();
                    for (int c = 0; c < chunk.ColCount; c++) headers.Add($"Column{c + 1}");
                }
            }

            // Build structured rows and optional markdown. Exclude header row on the first chunk when HeadersInFirstRow.
            int startRowOffset = (extract.HeadersInFirstRow && chunk.StartRow == r1) ? 1 : 0;
            var rowList = new List<IReadOnlyList<string>>(capacity: Math.Max(0, chunk.RowCount - startRowOffset));
            for (int r = startRowOffset; r < chunk.RowCount; r++) {
                var row = chunk.Rows[r];
                var cells = new string[chunk.ColCount];
                for (int c = 0; c < chunk.ColCount; c++) cells[c] = ToCellString(row[c], fallback: string.Empty);
                rowList.Add(cells);
            }

            bool truncated = false;
            int totalRows = rowList.Count;
            if (rowList.Count > chunking.MaxTableRows) {
                // Avoid duplicating memory by allocating a second list for large ranges.
                rowList.RemoveRange(chunking.MaxTableRows, rowList.Count - chunking.MaxTableRows);
                truncated = true;
            }

            string title = $"{sheetName} {range}";
            string[] columns = headers.ToArray();
            string? truncationWarning = truncated ? "Table rows truncated to MaxTableRows." : null;
            if (!extract.EmitMarkdownTable) {
                var table = new ExcelExtractTable {
                    Title = title,
                    Columns = columns,
                    Rows = rowList,
                    TotalRowCount = totalRows,
                    Truncated = truncated
                };
                yield return CreateChunk(sourcePath, sheetName, range, chunkIndex++, chunk.StartRow, RenderPlainTable(table), null, table,
                    truncationWarning == null ? null : new[] { truncationWarning });
                continue;
            }

            // Markdown is split between rows so every row stays in some chunk; the header repeats in each part.
            // A single row longer than MaxChars is kept whole in its own chunk.
            int firstSourceRow = chunk.StartRow + startRowOffset;
            int partStart = 0;
            bool firstPart = true;
            while (firstPart || partStart < rowList.Count) {
                int partEnd = MeasureMarkdownPart(title, columns, rowList, partStart, chunking.MaxChars, out bool oversizedRow);
                bool lastPart = partEnd >= rowList.Count;
                List<IReadOnlyList<string>> partRows = partStart == 0 && lastPart
                    ? rowList
                    : rowList.GetRange(partStart, partEnd - partStart);
                var table = new ExcelExtractTable {
                    Title = title,
                    Columns = columns,
                    Rows = partRows,
                    // The last part also accounts for rows MaxTableRows dropped from the block.
                    TotalRowCount = lastPart ? partRows.Count + (totalRows - rowList.Count) : partRows.Count,
                    Truncated = truncated && lastPart
                };
                string md = RenderMarkdownTable(table);
                var warnings = new List<string>(2);
                if (truncationWarning != null && lastPart) warnings.Add(truncationWarning);
                if (oversizedRow) warnings.Add("A table row with its header exceeds MaxChars; the row was kept whole in its own chunk.");
                yield return CreateChunk(sourcePath, sheetName, range, chunkIndex++, firstPart ? chunk.StartRow : firstSourceRow + partStart,
                    md, md, table, warnings.Count == 0 ? null : warnings.ToArray());
                firstPart = false;
                partStart = partEnd;
            }
        }
    }

    private static ExcelExtractChunk CreateChunk(string? sourcePath, string sheetName, string range, int chunkIndex, int startRow,
        string text, string? markdown, ExcelExtractTable table, string[]? warnings) =>
        new ExcelExtractChunk {
            Id = BuildStableId("excel", sourcePath, sheetName, chunkIndex, startRow),
            Location = new ExcelExtractLocation {
                Path = sourcePath,
                Sheet = sheetName,
                A1Range = range,
                BlockIndex = chunkIndex
            },
            Text = text,
            Markdown = markdown,
            Tables = new[] { table },
            Warnings = warnings
        };

    /// <summary>
    /// Returns the end (exclusive) of the longest run of rows from <paramref name="start"/> whose Markdown table,
    /// including the title and header, fits in <paramref name="maxChars"/>; at least one row is always taken.
    /// </summary>
    private static int MeasureMarkdownPart(string title, IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyList<string>> rows,
        int start, int maxChars, out bool oversizedRow) {
        oversizedRow = false;
        int newLine = Environment.NewLine.Length;
        int length = "### ".Length + title.Length + newLine * 2;
        length += 1 + newLine;
        foreach (string header in columns) length += EscapedLength(header) + 3;
        length += 1 + newLine + columns.Count * " --- |".Length;
        int index = start;
        while (index < rows.Count) {
            IReadOnlyList<string> row = rows[index];
            int rowLength = 1 + newLine;
            for (int column = 0; column < columns.Count; column++) {
                rowLength += EscapedLength(column < row.Count ? row[column] : string.Empty) + 3;
            }
            if (index > start && length + rowLength > maxChars) break;
            if (index == start && length + rowLength > maxChars) oversizedRow = true;
            length += rowLength;
            index++;
        }
        return index;
    }

    private static int EscapedLength(string? value) {
        if (string.IsNullOrEmpty(value)) return 0;
        int length = value!.Length;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '|') length++;
            else if (value[index] == (char)13 && index + 1 < value.Length && value[index + 1] == (char)10) { length--; index++; }
        }
        return length;
    }

    private static string ToCellString(object? value, string fallback) {
        if (value == null) return fallback;
        if (value is DateTime dt) return dt.ToString("o", CultureInfo.InvariantCulture);
        if (value is DateTimeOffset dto) return dto.ToString("o", CultureInfo.InvariantCulture);
        if (value is IFormattable f) return f.ToString(null, CultureInfo.InvariantCulture) ?? fallback;
        return value.ToString() ?? fallback;
    }

    private static string RenderPlainTable(ExcelExtractTable table) {
        // Simple human-readable fallback.
        var sb = new StringBuilder();
        sb.AppendLine(table.Title ?? "Table");
        sb.AppendLine(string.Join("\t", table.Columns));
        foreach (var row in table.Rows) sb.AppendLine(string.Join("\t", row));
        if (table.Truncated) sb.AppendLine("[truncated]");
        return sb.ToString().TrimEnd();
    }

    private static string RenderMarkdownTable(ExcelExtractTable table) {
        static string Esc(string s) => (s ?? string.Empty).Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ').Replace("|", "\\|");

        var sb = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(table.Title)) {
            sb.Append("### ").AppendLine(table.Title);
            sb.AppendLine();
        }

        sb.Append('|');
        foreach (var h in table.Columns) sb.Append(' ').Append(Esc(h)).Append(" |");
        sb.AppendLine();

        sb.Append('|');
        for (int i = 0; i < table.Columns.Count; i++) sb.Append(" --- |");
        sb.AppendLine();

        foreach (var row in table.Rows) {
            sb.Append('|');
            for (int i = 0; i < table.Columns.Count; i++) {
                var cell = (i < row.Count ? row[i] : string.Empty) ?? string.Empty;
                sb.Append(' ').Append(Esc(cell)).Append(" |");
            }
            sb.AppendLine();
        }

        if (table.Truncated) {
            sb.AppendLine();
            sb.AppendLine("<!-- truncated -->");
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildStableId(string kind, string? path, string sheet, int chunkIndex, int startRow) {
        var safe = string.IsNullOrWhiteSpace(path) ? "memory" : System.IO.Path.GetFileName(path!.Trim());
        var s = sheet.Trim().Replace(' ', '_');
        return $"{kind}:{safe}:{s}:c{chunkIndex}:r{startRow}";
    }
}
