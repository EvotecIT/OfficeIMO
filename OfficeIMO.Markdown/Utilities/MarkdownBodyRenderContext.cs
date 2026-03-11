namespace OfficeIMO.Markdown;

internal sealed class MarkdownBodyRenderContext {
    private readonly Dictionary<IMarkdownBlock, int> _blockIndexes;

    internal MarkdownBodyRenderContext(
        IReadOnlyList<IMarkdownBlock> blocks,
        HtmlOptions options,
        IReadOnlyDictionary<HeadingBlock, string> headingSlugs,
        MarkdownHeadingCatalog headingCatalog) {
        Blocks = blocks;
        Options = options;
        HeadingSlugs = headingSlugs;
        HeadingCatalog = headingCatalog;
        _blockIndexes = new Dictionary<IMarkdownBlock, int>(blocks.Count);
        for (int i = 0; i < blocks.Count; i++) {
            _blockIndexes[blocks[i]] = i;
        }
    }

    internal IReadOnlyList<IMarkdownBlock> Blocks { get; }
    internal HtmlOptions Options { get; }
    internal IReadOnlyDictionary<HeadingBlock, string> HeadingSlugs { get; }
    internal MarkdownHeadingCatalog HeadingCatalog { get; }

    internal int GetBlockIndex(IMarkdownBlock block) {
        if (block != null && _blockIndexes.TryGetValue(block, out var index)) {
            return index;
        }

        return -1;
    }
}
