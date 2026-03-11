namespace OfficeIMO.Markdown;

internal sealed class MarkdownBodyRenderContext {
    internal MarkdownBodyRenderContext(
        IReadOnlyList<IMarkdownBlock> blocks,
        HtmlOptions options,
        MarkdownHeadingCatalog headingCatalog) {
        Blocks = blocks;
        Options = options;
        HeadingCatalog = headingCatalog;
    }

    internal IReadOnlyList<IMarkdownBlock> Blocks { get; }
    internal HtmlOptions Options { get; }
    internal MarkdownHeadingCatalog HeadingCatalog { get; }
}
