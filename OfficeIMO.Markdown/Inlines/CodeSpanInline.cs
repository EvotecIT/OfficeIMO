namespace OfficeIMO.Markdown;

/// <summary>
/// Inline code span.
/// </summary>
public sealed class CodeSpanInline {
    /// <summary>Code content.</summary>
    public string Text { get; }
    /// <summary>Creates an inline code span.</summary>
    public CodeSpanInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() {
        // Choose a fence with length > any run of backticks in the text.
        int maxRun = 0; int run = 0;
        foreach (char c in Text) { if (c == '`') { run++; if (run > maxRun) maxRun = run; } else run = 0; }
        string fence = new string('`', maxRun + 1);
        // Per CommonMark, add a space inside when the text starts/ends with a backtick or space
        string leftPad = (Text.StartsWith("`") || Text.StartsWith(" ")) ? " " : string.Empty;
        string rightPad = (Text.EndsWith("`") || Text.EndsWith(" ")) ? " " : string.Empty;
        return fence + leftPad + Text + rightPad + fence;
    }
    internal string RenderHtml() => "<code>" + System.Net.WebUtility.HtmlEncode(Text) + "</code>";
}
