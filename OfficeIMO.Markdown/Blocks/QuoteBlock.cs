using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Simple blockquote block consisting of raw text lines.
/// </summary>
public sealed class QuoteBlock : IMarkdownBlock {
    public System.Collections.Generic.List<string> Lines { get; } = new System.Collections.Generic.List<string>();
    public System.Collections.Generic.List<IMarkdownBlock> Children { get; } = new System.Collections.Generic.List<IMarkdownBlock>();
    public QuoteBlock() { }
    public QuoteBlock(System.Collections.Generic.IEnumerable<string> lines) { Lines.AddRange(lines); }

    string IMarkdownBlock.RenderMarkdown() {
        if (Children.Count > 0) {
            var sb = new StringBuilder();
            for (int i = 0; i < Children.Count; i++) {
                var rendered = Children[i].RenderMarkdown();
                // Prefix every line with "> "
                using var reader = new System.IO.StringReader(rendered);
                string? line; bool first = true;
                while ((line = reader.ReadLine()) != null) {
                    if (!first) sb.AppendLine();
                    sb.Append("> ").Append(line);
                    first = false;
                }
                if (i < Children.Count - 1) sb.AppendLine().AppendLine("> "); // blank quote line to separate blocks
            }
            return sb.ToString();
        }
        var sb2 = new StringBuilder();
        foreach (var l in Lines) sb2.AppendLine("> " + l);
        return sb2.ToString().TrimEnd();
    }

    string IMarkdownBlock.RenderHtml() {
        if (Children.Count > 0) {
            var sb = new StringBuilder();
            sb.Append("<blockquote>");
            foreach (var b in Children) sb.Append(b.RenderHtml());
            sb.Append("</blockquote>");
            return sb.ToString();
        }
        var encoded = System.Net.WebUtility.HtmlEncode(string.Join("\n", Lines)).Replace("\n", "<br/>");
        return $"<blockquote><p>{encoded}</p></blockquote>";
    }
}
