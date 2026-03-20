namespace OfficeIMO.Markdown;

internal static class MarkdownDocumentBlockRewriter {
    public static void RewriteDocument(MarkdownDoc document, Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (blockRewriter == null) {
            throw new ArgumentNullException(nameof(blockRewriter));
        }

        document.Rewrite(new DelegateMarkdownRewriter(blockRewriter));
    }

    private sealed class DelegateMarkdownRewriter : MarkdownRewriter {
        private readonly Func<IMarkdownBlock, IMarkdownBlock> _blockRewriter;

        public DelegateMarkdownRewriter(Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
            _blockRewriter = blockRewriter ?? throw new ArgumentNullException(nameof(blockRewriter));
        }

        protected override IMarkdownBlock RewriteCurrentBlock(IMarkdownBlock block) =>
            _blockRewriter(block) ?? block;
    }
}
