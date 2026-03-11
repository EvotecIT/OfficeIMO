namespace OfficeIMO.Markdown;

internal interface IInlineContainerMarkdownInline {
    InlineSequence? NestedInlines { get; }
}
