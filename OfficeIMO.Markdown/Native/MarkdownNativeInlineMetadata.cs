namespace OfficeIMO.Markdown;

/// <summary>
/// Source-backed metadata attached to a native inline, such as a link target or image title.
/// </summary>
public sealed class MarkdownNativeInlineMetadata {
    internal MarkdownNativeInlineMetadata(string name, string value, MarkdownSyntaxNode syntaxNode)
        : this(name, value, syntaxNode, syntaxNode?.SourceSpan) {
    }

    internal MarkdownNativeInlineMetadata(string name, string value, MarkdownSyntaxNode syntaxNode, MarkdownSourceSpan? sourceSpan) {
        Name = name ?? string.Empty;
        Value = value ?? string.Empty;
        SyntaxNode = syntaxNode ?? throw new ArgumentNullException(nameof(syntaxNode));
        SourceSpan = sourceSpan;
    }

    /// <summary>Stable metadata name such as <c>target</c>, <c>title</c>, <c>alt</c>, or <c>source</c>.</summary>
    public string Name { get; }

    /// <summary>Metadata value.</summary>
    public string Value { get; }

    /// <summary>Syntax node that produced this metadata value.</summary>
    public MarkdownSyntaxNode SyntaxNode { get; }

    /// <summary>Source span for the metadata value when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }
}
