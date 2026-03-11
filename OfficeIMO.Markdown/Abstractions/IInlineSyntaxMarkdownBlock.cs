namespace OfficeIMO.Markdown;

internal interface IInlineSyntaxMarkdownBlock {
    InlineSequence SyntaxInlines { get; }
    MarkdownSyntaxKind SyntaxKind { get; }
    MarkdownSourceSpan? ProvidedSyntaxSpan { get; }
}
