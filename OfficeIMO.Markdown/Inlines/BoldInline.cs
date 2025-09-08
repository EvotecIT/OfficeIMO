namespace OfficeIMO.Markdown;

/// <summary>
/// Bold/strong emphasis inline.
/// </summary>
public sealed class BoldInline {
    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Creates a bold inline with the given text.</summary>
    public BoldInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => "**" + Text.Replace("**", "\\**") + "**";
    internal string RenderHtml() => "<strong>" + System.Net.WebUtility.HtmlEncode(Text) + "</strong>";
}
