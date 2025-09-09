using System.Collections.Generic;
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
        if (Headers.Count > 0) {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("| " + string.Join(" | ", Headers) + " |");
            var alignRow = new List<string>();
            for (int i = 0; i < Headers.Count; i++) {
                var a = (i < Alignments.Count) ? Alignments[i] : ColumnAlignment.None;
                alignRow.Add(a switch { ColumnAlignment.Left => ":---", ColumnAlignment.Center => ":---:", ColumnAlignment.Right => "---:", _ => "---" });
            }
            sb.AppendLine("| " + string.Join(" | ", alignRow) + " |");
            foreach (IReadOnlyList<string> row in Rows) sb.AppendLine("| " + string.Join(" | ", row) + " |");
            return sb.ToString().TrimEnd();
        } else {
            StringBuilder sb = new StringBuilder();
            foreach (IReadOnlyList<string> row in Rows) sb.AppendLine("| " + string.Join(" | ", row) + " |");
            return sb.ToString().TrimEnd();
        }
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
            sb.Append("<tr>");
            for (int i = 0; i < row.Count; i++) {
                var cell = row[i];
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
        // Minimal inline Markdown link recognizer: [text](url "opt title")
        // No nested brackets/parentheses handling; good enough for our generated links.
        if (cell.Length > 4 && cell[0] == '[') {
            int closeText = cell.IndexOf(']');
            if (closeText > 1 && closeText + 1 < cell.Length && cell[closeText + 1] == '(' && cell[cell.Length - 1] == ')') {
                string text = cell.Substring(1, closeText - 1);
                string inner = cell.Substring(closeText + 2, cell.Length - (closeText + 2) - 1);
                string url; string? title = null;
                int quoteIdx = inner.IndexOf('"');
                if (quoteIdx >= 0) {
                    // url "title"
                    url = inner.Substring(0, quoteIdx).Trim();
                    int quoteEnd = inner.LastIndexOf('"');
                    if (quoteEnd > quoteIdx) title = inner.Substring(quoteIdx + 1, quoteEnd - quoteIdx - 1);
                } else {
                    url = inner.Trim();
                }
                string titleAttr = string.IsNullOrEmpty(title) ? string.Empty : $" title=\"{System.Net.WebUtility.HtmlEncode(title)}\"";
                return $"<a href=\"{System.Net.WebUtility.HtmlEncode(url)}\"{titleAttr}>{System.Net.WebUtility.HtmlEncode(text)}</a>";
            }
        }
        return System.Net.WebUtility.HtmlEncode(cell);
    }
}
