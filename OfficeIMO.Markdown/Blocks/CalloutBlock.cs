using System.Text;

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

    /// <summary>Creates a callout with the specified kind, title and body.</summary>
    public CalloutBlock(string kind, string title, string body) {
        Kind = (kind ?? "info").Trim();
        Title = title ?? string.Empty;
        Body = body ?? string.Empty;
    }

    /// <inheritdoc />
    public string RenderMarkdown() {
        string tag = Kind.ToUpperInvariant();
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"> [!{tag}] {Title}");
        foreach (string line in Body.Replace("\r\n", "\n").Split('\n')) {
            sb.AppendLine("> " + line);
        }
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    public string RenderHtml() => $"<blockquote class=\"callout {System.Net.WebUtility.HtmlEncode(Kind)}\"><p><strong>{System.Net.WebUtility.HtmlEncode(Title)}</strong></p><p>{System.Net.WebUtility.HtmlEncode(Body)}</p></blockquote>";
}
