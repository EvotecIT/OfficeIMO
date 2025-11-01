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

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        static void AppendRow(StringBuilder builder, IEnumerable<string> cells) {
            builder.Append("| ");
            builder.Append(string.Join(" | ", cells));
            builder.Append(" |\n");
        }

        if (Headers.Count > 0) {
            var sb = new StringBuilder();
            var escapedHeaders = Headers.Select(EscapeMarkdownCell).ToList();
            AppendRow(sb, escapedHeaders);

            var alignRow = new List<string>();
            for (int i = 0; i < Headers.Count; i++) {
                var a = (i < Alignments.Count) ? Alignments[i] : ColumnAlignment.None;
                alignRow.Add(a switch { ColumnAlignment.Left => ":---", ColumnAlignment.Center => ":---:", ColumnAlignment.Right => "---:", _ => "---" });
            }
            AppendRow(sb, alignRow);

            foreach (IReadOnlyList<string> row in Rows) {
                var escapedRow = (row ?? Array.Empty<string>()).Select(EscapeMarkdownCell);
                AppendRow(sb, escapedRow);
            }

            return sb.ToString().TrimEnd('\n');
        }

        var sbNoHeaders = new StringBuilder();
        foreach (IReadOnlyList<string> row in Rows) {
            var escapedRow = (row ?? Array.Empty<string>()).Select(EscapeMarkdownCell);
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
            for (int i = 0; i < Headers.Count; i++) {
                var h = Headers[i];
                var style = (i < Alignments.Count) ? Alignments[i] : ColumnAlignment.None;
                var styleAttr = style switch { ColumnAlignment.Left => " style=\"text-align:left\"", ColumnAlignment.Center => " style=\"text-align:center\"", ColumnAlignment.Right => " style=\"text-align:right\"", _ => string.Empty };
                sb.Append($"<th{styleAttr}>{System.Net.WebUtility.HtmlEncode(h)}</th>");
            }
            sb.Append("</tr></thead>");
        }
        sb.Append("<tbody>");
        foreach (IReadOnlyList<string> row in Rows) {
            var cells = row ?? Array.Empty<string>();
            sb.Append("<tr>");
            for (int i = 0; i < cells.Count; i++) {
                var cell = cells[i] ?? string.Empty;
                var style = (i < Alignments.Count) ? Alignments[i] : ColumnAlignment.None;
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
        var inlines = MarkdownReader.ParseInlineText(normalized);
        var rendered = inlines.RenderHtml();
        return rendered.Contains('\n') ? rendered.Replace("\n", "<br/>") : rendered;
    }

    private static string EscapeMarkdownCell(string? cell) {
        if (string.IsNullOrEmpty(cell)) return string.Empty;

        var value = cell!;
        var builder = new StringBuilder(value.Length);

        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '\r':
                    if (i + 1 < value.Length && value[i + 1] == '\n') {
                        i++;
                    }
                    builder.Append("<br>");
                    break;
                case '\n':
                    builder.Append("<br>");
                    break;
                case '\\':
                    builder.Append("\\\\");
                    break;
                case '|':
                    builder.Append("\\|");
                    break;
                default:
                    builder.Append(ch);
                    break;
            }
        }

        return builder.ToString();
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
