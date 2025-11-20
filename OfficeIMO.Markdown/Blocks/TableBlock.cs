using System;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Pipe table with optional header row.
/// </summary>
public sealed class TableBlock : IMarkdownBlock {
    /// <summary>Optional header cells.</summary>
    public List<string> Headers { get; } = new List<string>();
    /// <summary>Data rows.</summary>
    public List<IReadOnlyList<string>> Rows { get; } = new List<IReadOnlyList<string>>();
    /// <summary>Optional column alignments per column (used when headers are present).</summary>
    public List<ColumnAlignment> Alignments { get; } = new List<ColumnAlignment>();
    /// <summary>Number of rows skipped due to table limits.</summary>
    public int SkippedRowCount { get; internal set; }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        static void AppendRow(StringBuilder builder, IReadOnlyList<string> cells) {
            builder.Append("| ");
            builder.Append(string.Join(" | ", cells));
            builder.Append(" |\n");
        }

        int columnCount = GetEffectiveColumnCount();

        if (Headers.Count > 0) {
            var sb = new StringBuilder();
            var preparedHeaders = PrepareRowCells(Headers, columnCount);
            var escapedHeaders = preparedHeaders.Select(EscapeMarkdownCell).ToArray();
            AppendRow(sb, escapedHeaders);

            var alignRow = new string[columnCount];
            for (int i = 0; i < columnCount; i++) {
                var a = GetAlignment(i);
                alignRow[i] = a switch { ColumnAlignment.Left => ":---", ColumnAlignment.Center => ":---:", ColumnAlignment.Right => "---:", _ => "---" };
            }
            AppendRow(sb, alignRow);

            foreach (IReadOnlyList<string> row in Rows) {
                var preparedRow = PrepareRowCells(row, columnCount);
                var escapedRow = preparedRow.Select(EscapeMarkdownCell).ToArray();
                AppendRow(sb, escapedRow);
            }

            return sb.ToString().TrimEnd('\n');
        }

        var sbNoHeaders = new StringBuilder();
        foreach (IReadOnlyList<string> row in Rows) {
            var preparedRow = PrepareRowCells(row, columnCount);
            var escapedRow = preparedRow.Select(EscapeMarkdownCell).ToArray();
            AppendRow(sbNoHeaders, escapedRow);
        }
        return sbNoHeaders.ToString().TrimEnd('\n');
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        StringBuilder sb = new StringBuilder();
        sb.Append("<table>");
        if (Headers.Count > 0) {
            sb.Append("<thead><tr>");
            int columnCount = GetEffectiveColumnCount();
            var preparedHeaders = PrepareRowCells(Headers, columnCount);
            for (int i = 0; i < preparedHeaders.Count; i++) {
                var h = preparedHeaders[i];
                var style = GetAlignment(i);
                var styleAttr = style switch { ColumnAlignment.Left => " style=\"text-align:left\"", ColumnAlignment.Center => " style=\"text-align:center\"", ColumnAlignment.Right => " style=\"text-align:right\"", _ => string.Empty };
                sb.Append($"<th{styleAttr}>{System.Net.WebUtility.HtmlEncode(h)}</th>");
            }
            sb.Append("</tr></thead>");
        }
        sb.Append("<tbody>");
        int bodyColumnCount = GetEffectiveColumnCount();
        foreach (IReadOnlyList<string> row in Rows) {
            var cells = PrepareRowCells(row, bodyColumnCount);
            sb.Append("<tr>");
            for (int i = 0; i < cells.Count; i++) {
                var cell = cells[i];
                var style = GetAlignment(i);
                var styleAttr = style switch { ColumnAlignment.Left => " style=\"text-align:left\"", ColumnAlignment.Center => " style=\"text-align:center\"", ColumnAlignment.Right => " style=\"text-align:right\"", _ => string.Empty };
                sb.Append($"<td{styleAttr}>");
                sb.Append(RenderCellHtml(cell));
                sb.Append("</td>");
            }
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table>");
        return sb.ToString();
    }

    private static string RenderCellHtml(string cell) {
        if (string.IsNullOrEmpty(cell)) return string.Empty;
        // Allow simple <br> markers inside table cells and support inline markdown (code, links, emphasis).
        // We avoid allowing arbitrary HTML by translating only <br> tags to hard breaks and then using the inline parser.
        var normalized = NormalizeBreakMarkers(cell);
        var sanitized = SanitizeInlineMarkdownInput(normalized);
        var inlines = MarkdownReader.ParseInlineText(sanitized);
        var rendered = inlines.RenderHtml();
        rendered = NormalizeEncodedEntities(rendered);
        return rendered.Contains('\n') ? rendered.Replace("\n", "<br/>") : rendered;
    }

    private static string EscapeMarkdownCell(string? cell) {
        if (string.IsNullOrEmpty(cell)) return string.Empty;

        var value = cell!;
        StringBuilder? builder = null;

        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '\r':
                    builder ??= AllocateCellBuilder(value, i);
                    if (i + 1 < value.Length && value[i + 1] == '\n') {
                        i++;
                    }
                    builder.Append("<br>");
                    break;
                case '\n':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("<br>");
                    break;
                case '\\':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("\\\\");
                    break;
                case '|':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("\\|");
                    break;
                default:
                    builder?.Append(ch);
                    break;
            }
        }

        return builder?.ToString() ?? value;
    }

    private static string NormalizeBreakMarkers(string cell) {
        var builder = new StringBuilder(cell.Length);

        for (int i = 0; i < cell.Length; i++) {
            char ch = cell[i];
            switch (ch) {
                case '\r':
                    if (i + 1 < cell.Length && cell[i + 1] == '\n') {
                        i++;
                    }
                    builder.Append('\n');
                    break;
                case '\n':
                    builder.Append('\n');
                    break;
                case '<':
                    if (TryConsumeBreakTag(cell, i, out int consumed)) {
                        builder.Append('\n');
                        i += consumed - 1;
                    } else {
                        builder.Append(ch);
                    }
                    break;
                default:
                    builder.Append(ch);
                    break;
            }
        }

        return builder.ToString();
    }

    private static string SanitizeInlineMarkdownInput(string value) {
        StringBuilder? builder = null;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '<':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("&lt;");
                    break;
                case '>':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("&gt;");
                    break;
                case '&':
                    builder ??= AllocateCellBuilder(value, i);
                    builder.Append("&amp;");
                    break;
                default:
                    builder?.Append(ch);
                    break;
            }
        }

        return builder?.ToString() ?? value;
    }

    private static string NormalizeEncodedEntities(string value) {
        if (value.IndexOf("&amp;", StringComparison.Ordinal) < 0) {
            return value;
        }

        return value
            .Replace("&amp;lt;", "&lt;")
            .Replace("&amp;gt;", "&gt;")
            .Replace("&amp;amp;", "&amp;");
    }

    private static StringBuilder AllocateCellBuilder(string seed, int copyLength) {
        var builder = new StringBuilder(seed.Length + 8);
        if (copyLength > 0) {
            builder.Append(seed, 0, copyLength);
        }
        return builder;
    }

    private int GetEffectiveColumnCount() {
        int columnCount = Headers.Count;

        foreach (var row in Rows) {
            if (row != null) {
                columnCount = Math.Max(columnCount, row.Count);
            }
        }

        columnCount = Math.Max(columnCount, Alignments.Count);
        return columnCount;
    }

    private ColumnAlignment GetAlignment(int index) {
        if (index < 0) return ColumnAlignment.None;
        return index < Alignments.Count ? Alignments[index] : ColumnAlignment.None;
    }

    private static IReadOnlyList<string> PrepareRowCells(IReadOnlyList<string>? row, int expectedCount) {
        if (row == null || row.Count == 0) {
            if (expectedCount <= 0) {
                return Array.Empty<string>();
            }

            var padded = new string[expectedCount];
            for (int i = 0; i < padded.Length; i++) {
                padded[i] = string.Empty;
            }
            return padded;
        }

        if (expectedCount <= 0) {
            var copy = new string[row.Count];
            for (int i = 0; i < row.Count; i++) {
                copy[i] = row[i] ?? string.Empty;
            }
            return copy;
        }

        var cells = new string[expectedCount];
        int limit = Math.Min(expectedCount, row.Count);
        for (int i = 0; i < limit; i++) {
            cells[i] = row[i] ?? string.Empty;
        }
        if (limit < expectedCount) {
            for (int i = limit; i < expectedCount; i++) {
                cells[i] = string.Empty;
            }
        }
        return cells;
    }

    private static bool TryConsumeBreakTag(string value, int index, out int consumed) {
        consumed = 0;
        int length = value.Length;
        if (index + 3 >= length) {
            return false;
        }

        if (value[index] != '<') return false;
        if (!IsSpecificLetter(value[index + 1], 'b')) return false;
        if (!IsSpecificLetter(value[index + 2], 'r')) return false;

        int position = index + 3;

        while (position < length && char.IsWhiteSpace(value[position])) {
            position++;
        }

        if (position < length && value[position] == '/') {
            position++;
            while (position < length && char.IsWhiteSpace(value[position])) {
                position++;
            }
        }

        if (position < length && value[position] == '>') {
            consumed = position - index + 1;
            return true;
        }

        return false;
    }

    private static bool IsSpecificLetter(char value, char expected) {
        return char.ToLowerInvariant(value) == expected;
    }
}
