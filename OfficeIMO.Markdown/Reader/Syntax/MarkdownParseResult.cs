namespace OfficeIMO.Markdown;

/// <summary>
/// Result of parsing markdown into both the object model and a syntax tree.
/// </summary>
public sealed class MarkdownParseResult {
    /// <summary>The parsed markdown object model.</summary>
    public MarkdownDoc Document { get; }
    /// <summary>The root syntax node for the parsed markdown.</summary>
    public MarkdownSyntaxNode SyntaxTree { get; }

    internal MarkdownParseResult(MarkdownDoc document, MarkdownSyntaxNode syntaxTree) {
        Document = document;
        SyntaxTree = syntaxTree;
    }
}
