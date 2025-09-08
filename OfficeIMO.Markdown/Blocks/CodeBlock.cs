using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Fenced code block with optional caption.
/// </summary>
public sealed class CodeBlock : IMarkdownBlock, ICaptionable {
    /// <summary>Optional language hint (e.g., csharp, bash).</summary>
    public string Language { get; }
    /// <summary>Code contents.</summary>
    public string Content { get; }
    /// <summary>Optional caption shown under the block.</summary>
    public string? Caption { get; set; }

    /// <summary>Create a code block with a language hint.</summary>
    public CodeBlock(string language, string content) {
        Language = language ?? string.Empty;
        Content = content ?? string.Empty;
    }

    /// <inheritdoc />
    public string RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"```{Language}");
        sb.AppendLine(Content);
        sb.AppendLine("```");
        if (!string.IsNullOrWhiteSpace(Caption)) sb.AppendLine("_" + Caption + "_");
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    public string RenderHtml() {
        string lang = string.IsNullOrEmpty(Language) ? string.Empty : $" class=\"language-{System.Net.WebUtility.HtmlEncode(Language)}\"";
        string code = System.Net.WebUtility.HtmlEncode(Content);
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
        return $"<pre><code{lang}>{code}</code></pre>{caption}";
    }
}
