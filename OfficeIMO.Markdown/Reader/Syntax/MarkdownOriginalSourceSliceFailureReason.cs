namespace OfficeIMO.Markdown;

/// <summary>
/// Explains why an original-input source slice could not be materialized for a normalized source span.
/// </summary>
public enum MarkdownOriginalSourceSliceFailureReason {
    /// <summary>No failure occurred.</summary>
    None,
    /// <summary>The parse result did not retain exact original reader input.</summary>
    OriginalMarkdownNotPreserved,
    /// <summary>No final syntax node was found for the supplied associated object.</summary>
    AssociatedObjectNotFound,
    /// <summary>The requested syntax node or span has no source span that can be mapped.</summary>
    SourceSpanUnavailable,
    /// <summary>The requested syntax node was generated from semantic content and has no exact original source.</summary>
    GeneratedSyntaxNode,
    /// <summary>The original reader input is not equivalent to the normalized source text.</summary>
    OriginalTextNotEquivalent,
    /// <summary>The source span could not be mapped into the original reader input.</summary>
    OriginalSpanUnavailable
}
