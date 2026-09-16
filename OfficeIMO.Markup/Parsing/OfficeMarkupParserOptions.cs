using OfficeIMO.Markdown;

namespace OfficeIMO.Markup;

/// <summary>
/// Options for parsing Markdown-inspired OfficeIMO markup into the semantic AST.
/// </summary>
public sealed class OfficeMarkupParserOptions {
    /// <summary>Gets or sets the fallback authoring profile when front matter does not select one; defaults to Document.</summary>
    public OfficeMarkupProfile Profile { get; set; } = OfficeMarkupProfile.Document;
    /// <summary>Gets or sets whether profile and required-field validation diagnostics are added; defaults to true.</summary>
    public bool Validate { get; set; } = true;
    /// <summary>Gets or sets optional settings passed to the underlying Markdown semantic parser.</summary>
    public MarkdownReaderOptions? MarkdownOptions { get; set; }
}
