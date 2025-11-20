namespace OfficeIMO.Markdown;

/// <summary>
/// Fenced code block with optional caption. Fence length adapts to backticks inside the content.
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
    string IMarkdownBlock.RenderMarkdown() {
        // Choose a fence with length > any run of backticks in the content to avoid premature closure.
        int maxRun = 0;
        int run = 0;
        foreach (char c in Content) {
            if (c == '`') {
                run++;
                if (run > maxRun) maxRun = run;
            } else {
                run = 0;
            }
        }

        string fence = new string('`', Math.Max(3, maxRun + 1));

        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"{fence}{Language}");
        sb.AppendLine(Content);
        sb.AppendLine(fence);
        if (!string.IsNullOrWhiteSpace(Caption)) sb.AppendLine("_" + Caption + "_");
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        string lang = string.IsNullOrEmpty(Language) ? string.Empty : $" class=\"language-{System.Net.WebUtility.HtmlEncode(Language)}\"";
        string code = System.Net.WebUtility.HtmlEncode(Content);
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
        return $"<pre><code{lang}>{code}</code></pre>{caption}";
    }
}
