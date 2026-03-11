namespace OfficeIMO.Markdown;

internal interface ITocPlaceholderMarkdownBlock : IBodySidebarMarkdownBlock {
    bool RequiresScrollSpy();
    TocBlock RealizeToc(IReadOnlyList<IMarkdownBlock> blocks, int placeholderIndex, MarkdownHeadingCatalog headingCatalog);
}
