using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public sealed class TableBlock : IMarkdownBlock {
    public List<string> Headers { get; } = new List<string>();
    public List<IReadOnlyList<string>> Rows { get; } = new List<IReadOnlyList<string>>();

    public string RenderMarkdown() {
        if (Headers.Count > 0) {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("| " + string.Join(" | ", Headers) + " |");
            sb.AppendLine("| " + string.Join(" | ", Headers.Select(_ => "---")) + " |");
            foreach (IReadOnlyList<string> row in Rows) sb.AppendLine("| " + string.Join(" | ", row) + " |");
            return sb.ToString().TrimEnd();
        } else {
            StringBuilder sb = new StringBuilder();
            foreach (IReadOnlyList<string> row in Rows) sb.AppendLine("| " + string.Join(" | ", row) + " |");
            return sb.ToString().TrimEnd();
        }
    }

    public string RenderHtml() {
        StringBuilder sb = new StringBuilder();
        sb.Append("<table>");
        if (Headers.Count > 0) {
            sb.Append("<thead><tr>");
            foreach (string h in Headers) sb.Append($"<th>{System.Net.WebUtility.HtmlEncode(h)}</th>");
            sb.Append("</tr></thead>");
        }
        sb.Append("<tbody>");
        foreach (IReadOnlyList<string> row in Rows) {
            sb.Append("<tr>");
            foreach (string cell in row) sb.Append($"<td>{System.Net.WebUtility.HtmlEncode(cell)}</td>");
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table>");
        return sb.ToString();
    }
}

