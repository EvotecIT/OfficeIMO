using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel;

/// <summary>
/// Chunked extraction helpers intended for AI ingestion.
/// </summary>
public static class ExcelExtractionExtensions {
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
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentNullException(nameof(sheetName));

        extract ??= new ExcelExtractOptions();
        chunking ??= new ExcelExtractChunkingOptions();
        if (extract.ChunkRows < 1) extract.ChunkRows = 1;
        if (chunking.MaxChars < 256) chunking.MaxChars = 256;
        if (chunking.MaxTableRows < 1) chunking.MaxTableRows = 1;

        var sheet = reader.GetSheet(sheetName);
        var range = string.IsNullOrWhiteSpace(a1Range) ? sheet.GetUsedRangeA1() : a1Range!;

        // Determine headers from the first row in the overall range.
        (int r1, int c1, int r2, int c2) = A1.ParseRange(range);
        var width = Math.Max(0, c2 - c1 + 1);
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
                rowList = rowList.Take(chunking.MaxTableRows).ToList();
                truncated = true;
            }

            var table = new ExcelExtractTable {
                Title = $"{sheetName} {range}",
                Columns = headers.ToArray(),
                Rows = rowList,
                TotalRowCount = totalRows,
                Truncated = truncated
            };

            string? md = null;
            if (extract.EmitMarkdownTable) {
                md = RenderMarkdownTable(table);
                if (md.Length > chunking.MaxChars) {
                    md = md.Substring(0, chunking.MaxChars) + "\n\n<!-- truncated -->";
                }
            }

            var id = BuildStableId("excel", sourcePath, sheetName, chunkIndex, chunk.StartRow);
            yield return new ExcelExtractChunk {
                Id = id,
                Location = new ExcelExtractLocation {
                    Path = sourcePath,
                    Sheet = sheetName,
                    A1Range = range,
                    BlockIndex = chunkIndex
                },
                Text = md ?? RenderPlainTable(table),
                Markdown = md,
                Tables = new[] { table },
                Warnings = truncated ? new[] { "Table rows truncated to MaxTableRows." } : null
            };

            chunkIndex++;
        }
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

