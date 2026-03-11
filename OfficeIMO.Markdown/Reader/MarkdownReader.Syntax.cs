namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static MarkdownSyntaxNode BuildDocumentSyntaxTree(IReadOnlyList<MarkdownSyntaxNode> children) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.Document, MarkdownBlockSyntaxBuilder.GetAggregateSpan(children), children: children);

    private static void CaptureSyntaxNodes(MarkdownDoc doc, int previousBlockCount, int startLine, int endExclusiveLine, List<MarkdownSyntaxNode> nodes) {
        int start = startLine + 1;
        int end = Math.Max(start, endExclusiveLine);
        var span = new MarkdownSourceSpan(start, end);

        for (int blockIndex = previousBlockCount; blockIndex < doc.Blocks.Count; blockIndex++) {
            nodes.Add(BuildSyntaxNode(doc.Blocks[blockIndex], span));
        }
    }

    private static MarkdownSyntaxNode BuildSyntaxNode(IMarkdownBlock block, MarkdownSourceSpan? span = null) =>
        MarkdownBlockSyntaxBuilder.BuildBlock(block, span);
}
