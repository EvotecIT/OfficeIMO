namespace OfficeIMO.Markdown;

/// <summary>
/// Describes how a normalized markdown source span maps back to the preserved original reader input.
/// </summary>
public enum MarkdownOriginalSourceMappingKind {
    /// <summary>The span could not be mapped to original reader input.</summary>
    Unavailable,
    /// <summary>The original reader input and normalized markdown text are byte-identical for the document.</summary>
    Exact,
    /// <summary>The original reader input differs only by line-ending spelling and the span was mapped across that difference.</summary>
    LineEndingEquivalent
}
