namespace OfficeIMO.Markdown;

/// <summary>
/// Docs/Markdown-style callout (admonition) block. Renders using
/// "> [!KIND] Title" followed by indented content lines.
/// </summary>
public sealed class CalloutBlock : IMarkdownBlock {
    /// <summary>Admonition kind, e.g., info, warning, success.</summary>
    public string Kind { get; }
    /// <summary>Callout title displayed inline with the marker.</summary>
    public string Title { get; }
    /// <summary>Callout body text (can include multiple lines).</summary>
    public string Body { get; }

    /// <summary>
    /// Optional parsed body blocks. When present (produced by <see cref="MarkdownReader"/>),
    /// HTML/Markdown rendering uses these blocks instead of the raw <see cref="Body"/> string.
    /// </summary>
    internal IReadOnlyList<IMarkdownBlock>? Children { get; }

    /// <summary>Creates a callout with the specified kind, title and body.</summary>
    public CalloutBlock(string kind, string title, string body) {
        Kind = (kind ?? "info").Trim();
        Title = title ?? string.Empty;
        Body = body ?? string.Empty;
    }

    internal CalloutBlock(string kind, string title, IReadOnlyList<IMarkdownBlock> children) {
        Kind = (kind ?? "info").Trim();
        Title = title ?? string.Empty;
        Body = string.Empty;
        Children = children;
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        string tag = Kind.ToUpperInvariant();
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"> [!{tag}] {Title}");
        string bodyMarkdown;
        if (Children != null && Children.Count > 0) {
            var inner = new StringBuilder();
            for (int i = 0; i < Children.Count; i++) {
                if (Children[i] == null) continue;
                var rendered = Children[i].RenderMarkdown();
                if (string.IsNullOrEmpty(rendered)) continue;
                inner.AppendLine(rendered.TrimEnd());
            }
            bodyMarkdown = inner.ToString().TrimEnd();
        } else {
            bodyMarkdown = Body ?? string.Empty;
        }
        foreach (string line in bodyMarkdown.Replace("\r\n", "\n").Split('\n')) {
            sb.AppendLine(line.Length == 0 ? ">" : ("> " + line));
        }
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var kind = System.Net.WebUtility.HtmlEncode(Kind);
        var title = System.Net.WebUtility.HtmlEncode(Title);

        var sb = new StringBuilder();
        sb.Append("<blockquote class=\"callout ").Append(kind).Append("\">");
        sb.Append("<p><strong>").Append(title).Append("</strong></p>");

        if (Children != null && Children.Count > 0) {
            for (int i = 0; i < Children.Count; i++) {
                if (Children[i] == null) continue;
                sb.Append(Children[i].RenderHtml());
            }
        } else {
            // Plain text body (builder-created callouts).
            var body = (Body ?? string.Empty).Replace("\r\n", "\n");
            var lines = body.Split('\n');
            sb.Append("<p>");
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) sb.Append("<br/>");
                sb.Append(System.Net.WebUtility.HtmlEncode(lines[i]));
            }
            sb.Append("</p>");
        }

        sb.Append("</blockquote>");
        return sb.ToString();
    }
}
