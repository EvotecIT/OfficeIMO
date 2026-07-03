namespace OfficeIMO.Markdown;

/// <summary>
/// Diagnostic describing a final syntax node that was generated from semantic content
/// instead of being parsed as an exact source-backed node.
/// </summary>
public sealed class MarkdownGeneratedSyntaxDiagnostic {
    internal MarkdownGeneratedSyntaxDiagnostic(MarkdownSyntaxNode syntaxNode, string syntaxPath, string indexPath) {
        SyntaxNode = syntaxNode ?? throw new ArgumentNullException(nameof(syntaxNode));
        SyntaxPath = syntaxPath ?? string.Empty;
        IndexPath = indexPath ?? string.Empty;
        SourceSpan = syntaxNode.SourceSpan;
        AssociatedObject = syntaxNode.AssociatedObject;
        AssociatedObjectType = syntaxNode.AssociatedObject?.GetType().Name;
    }

    /// <summary>Stable diagnostic identifier.</summary>
    public string Id => "syntax.generated-node";

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message => "Final syntax node was generated from semantic content; any source span is a fallback anchor, not exact parsed source.";

    /// <summary>The generated final syntax node.</summary>
    public MarkdownSyntaxNode SyntaxNode { get; }

    /// <summary>Generated syntax node kind.</summary>
    public MarkdownSyntaxKind Kind => SyntaxNode.Kind;

    /// <summary>Optional custom extension kind for generated extension nodes.</summary>
    public string? CustomKind => SyntaxNode.CustomKind;

    /// <summary>Optional generated-node source span. When present, this is a fallback anchor.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Path of syntax kinds from the document root to the generated node.</summary>
    public string SyntaxPath { get; }

    /// <summary>Zero-based child-index path from the document root to the generated node.</summary>
    public string IndexPath { get; }

    /// <summary>Semantic object associated with the generated syntax node when available.</summary>
    public object? AssociatedObject { get; }

    /// <summary>Semantic object type name associated with the generated syntax node when available.</summary>
    public string? AssociatedObjectType { get; }
}
