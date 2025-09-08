using System.Text;

namespace OfficeIMO.Markdown;

public sealed class CalloutBlock : IMarkdownBlock {
    public string Kind { get; }
    public string Title { get; }
    public string Body { get; }

    public CalloutBlock(string kind, string title, string body) {
        Kind = (kind ?? "info").Trim();
        Title = title ?? string.Empty;
        Body = body ?? string.Empty;
    }

    public string RenderMarkdown() {
        string tag = Kind.ToUpperInvariant();
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"> [!{tag}] {Title}");
        foreach (string line in Body.Replace("\r\n", "\n").Split('\n')) {
            sb.AppendLine("> " + line);
        }
        return sb.ToString().TrimEnd();
    }

    public string RenderHtml() => $"<blockquote class=\"callout {System.Net.WebUtility.HtmlEncode(Kind)}\"><p><strong>{System.Net.WebUtility.HtmlEncode(Title)}</strong></p><p>{System.Net.WebUtility.HtmlEncode(Body)}</p></blockquote>";
}

