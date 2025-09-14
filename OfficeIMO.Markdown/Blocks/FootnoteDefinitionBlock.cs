using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote definition block, e.g., [^id]: content.
/// </summary>
public sealed class FootnoteDefinitionBlock : IMarkdownBlock {
    public string Label { get; }
    public string Text { get; }
    public FootnoteDefinitionBlock(string label, string text) { Label = label ?? string.Empty; Text = text ?? string.Empty; }
    string IMarkdownBlock.RenderMarkdown() => $"[^{Label}]: {Text}";
    string IMarkdownBlock.RenderHtml() => $"<p id=\"fn:{System.Net.WebUtility.HtmlEncode(Label)}\"><sup>{System.Net.WebUtility.HtmlEncode(Label)}</sup> {System.Net.WebUtility.HtmlEncode(Text)}</p>";
}

