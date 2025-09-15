using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote definition block, e.g., [^id]: content.
/// </summary>
public sealed class FootnoteDefinitionBlock : IMarkdownBlock {
    /// <summary>Footnote label (identifier without the leading ^).</summary>
    public string Label { get; }
    /// <summary>Footnote text content.</summary>
    public string Text { get; }
    /// <summary>Create a new footnote definition.</summary>
    /// <param name="label">Identifier used by references.</param>
    /// <param name="text">Definition text.</param>
    public FootnoteDefinitionBlock(string label, string text) { Label = label ?? string.Empty; Text = text ?? string.Empty; }
    string IMarkdownBlock.RenderMarkdown() => $"[^{Label}]: {Text}";
    string IMarkdownBlock.RenderHtml() => $"<p id=\"fn:{System.Net.WebUtility.HtmlEncode(Label)}\"><sup>{System.Net.WebUtility.HtmlEncode(Label)}</sup> {System.Net.WebUtility.HtmlEncode(Text)}</p>";
}
