using OfficeIMO.Markdown;

namespace OfficeIMO.Markup;

/// <summary>
/// Options for parsing Markdown-inspired OfficeIMO markup into the semantic AST.
/// </summary>
public sealed class OfficeMarkupParserOptions {
    public OfficeMarkupProfile Profile { get; set; } = OfficeMarkupProfile.Document;
    public bool Validate { get; set; } = true;
    public MarkdownReaderOptions? MarkdownOptions { get; set; }
}
