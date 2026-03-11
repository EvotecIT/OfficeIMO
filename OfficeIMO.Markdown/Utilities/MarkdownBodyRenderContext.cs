namespace OfficeIMO.Markdown;

internal sealed class MarkdownBodyRenderContext {
    internal MarkdownBodyRenderContext(
        IReadOnlyList<IMarkdownBlock> blocks,
        HtmlOptions options,
        IReadOnlyDictionary<IHeadingMarkdownBlock, string> headingSlugs,
        MarkdownHeadingCatalog headingCatalog) {
        Blocks = blocks;
        Options = options;
        HeadingSlugs = headingSlugs;
        HeadingCatalog = headingCatalog;
    }

    internal IReadOnlyList<IMarkdownBlock> Blocks { get; }
    internal HtmlOptions Options { get; }
    internal IReadOnlyDictionary<IHeadingMarkdownBlock, string> HeadingSlugs { get; }
    internal MarkdownHeadingCatalog HeadingCatalog { get; }
}
