namespace OfficeIMO.Markdown;

internal static class MarkdownDocumentBlockRewriter {
    public static void RewriteDocument(MarkdownDocument document, Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        ArgumentNullException.ThrowIfNull(document);

        ArgumentNullException.ThrowIfNull(blockRewriter);

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
