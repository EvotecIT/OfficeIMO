using System.Text;

namespace OfficeIMO.Markdown;

public sealed class CodeBlock : IMarkdownBlock, ICaptionable {
    public string Language { get; }
    public string Content { get; }
    public string? Caption { get; set; }

    public CodeBlock(string language, string content) {
        Language = language ?? string.Empty;
        Content = content ?? string.Empty;
    }

    public string RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"```{Language}");
        sb.AppendLine(Content);
        sb.AppendLine("```");
        if (!string.IsNullOrWhiteSpace(Caption)) sb.AppendLine("_" + Caption + "_");
        return sb.ToString().TrimEnd();
    }

    public string RenderHtml() {
        string lang = string.IsNullOrEmpty(Language) ? string.Empty : $" class=\"language-{System.Net.WebUtility.HtmlEncode(Language)}\"";
        string code = System.Net.WebUtility.HtmlEncode(Content);
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
        return $"<pre><code{lang}>{code}</code></pre>{caption}";
    }
}

