namespace OfficeIMO.Markdown;

internal sealed class MarkdownRenderModel {
    internal MarkdownRenderModel(
        IReadOnlyList<IMarkdownBlock> sourceBlocks,
        IReadOnlyList<IMarkdownBlock> realizedBlocks,
        IReadOnlyDictionary<HeadingBlock, string> headingSlugs,
        MarkdownHeadingCatalog headingCatalog,
        bool hasScrollSpyToc) {
        SourceBlocks = sourceBlocks;
        RealizedBlocks = realizedBlocks;
        HeadingSlugs = headingSlugs;
        HeadingCatalog = headingCatalog;
        HasScrollSpyToc = hasScrollSpyToc;
    }

    internal IReadOnlyList<IMarkdownBlock> SourceBlocks { get; }
    internal IReadOnlyList<IMarkdownBlock> RealizedBlocks { get; }
    internal IReadOnlyDictionary<HeadingBlock, string> HeadingSlugs { get; }
    internal MarkdownHeadingCatalog HeadingCatalog { get; }
    internal bool HasScrollSpyToc { get; }
}
