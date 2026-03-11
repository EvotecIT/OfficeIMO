namespace OfficeIMO.Markdown;

internal interface IParagraphMarkdownBlock : ITightListItemHtmlMarkdownBlock {
    InlineSequence ParagraphInlines { get; }
}
