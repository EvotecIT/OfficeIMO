namespace OfficeIMO.Markdown;

public sealed class TextRun {
    public string Text { get; }
    public TextRun(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => Escape(Text);
    internal string RenderHtml() => System.Net.WebUtility.HtmlEncode(Text);
    private static string Escape(string s) {
        return s.Replace("\\", "\\\\").Replace("*", "\\*").Replace("_", "\\_");
    }
}

