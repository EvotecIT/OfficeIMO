namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static MarkdownSyntaxNode BuildDocumentSyntaxTree(IReadOnlyList<MarkdownSyntaxNode> children, MarkdownDoc? document = null) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.Document, MarkdownBlockSyntaxBuilder.GetAggregateSpan(children), children: children, associatedObject: document);

    internal static MarkdownSyntaxNode BuildSyntaxTree(
        MarkdownDoc document,
        IReadOnlyList<MarkdownSourceSpan?>? topLevelBlockSourceSpans = null,
        MarkdownSourceSpan? frontMatterSpan = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var children = new List<MarkdownSyntaxNode>(document.TopLevelBlocks.Count);
        if (document.DocumentHeader != null) {
            children.Add(BuildSyntaxNode(document.DocumentHeader, frontMatterSpan));
        }

        for (int i = 0; i < document.Blocks.Count; i++) {
            var span = topLevelBlockSourceSpans != null && i < topLevelBlockSourceSpans.Count
                ? topLevelBlockSourceSpans[i]
                : null;
            children.Add(BuildSyntaxNode(document.Blocks[i], span));
        }

        return BuildDocumentSyntaxTree(children, document);
    }

    internal static MarkdownSyntaxNode BuildFinalSyntaxTree(
        MarkdownDoc document,
        MarkdownSyntaxNode originalSyntaxTree,
        IReadOnlyList<MarkdownDocumentTransformDiagnostic>? transformDiagnostics = null) {
        var blockSpans = BuildTopLevelBlockSourceSpans(document, originalSyntaxTree, transformDiagnostics);
        var frontMatterSpan = GetFrontMatterSpan(document, originalSyntaxTree);
        return BuildSyntaxTree(document, blockSpans, frontMatterSpan);
    }

    private static void CaptureSyntaxNodes(MarkdownDoc doc, int previousBlockCount, int startLine, int endExclusiveLine, List<MarkdownSyntaxNode> nodes, MarkdownReaderState? state = null) {
        int start = startLine + 1;
        int end = Math.Max(start, endExclusiveLine);
        var span = CreateLineSpan(state, start, end);

        for (int blockIndex = previousBlockCount; blockIndex < doc.Blocks.Count; blockIndex++) {
            nodes.Add(BuildSyntaxNode(doc.Blocks[blockIndex], span));
        }
    }

    private static MarkdownSyntaxNode BuildSyntaxNode(IMarkdownBlock block, MarkdownSourceSpan? span = null) =>
        MarkdownBlockSyntaxBuilder.BuildBlock(block, span);

    private static MarkdownSourceSpan CreateLineSpan(MarkdownReaderState? state, int startLine, int endLine) =>
        state?.SourceTextMap?.CreateLineSpan(startLine, endLine) ?? new MarkdownSourceSpan(startLine, endLine);

    private static MarkdownSourceSpan CreateSpan(MarkdownReaderState? state, int startLine, int startColumn, int endLine, int endColumn) =>
        state?.SourceTextMap?.CreateSpan(startLine, startColumn, endLine, endColumn)
        ?? new MarkdownSourceSpan(startLine, startColumn, endLine, endColumn);

    private static IReadOnlyList<MarkdownSourceSpan?> BuildTopLevelBlockSourceSpans(
        MarkdownDoc document,
        MarkdownSyntaxNode originalSyntaxTree,
        IReadOnlyList<MarkdownDocumentTransformDiagnostic>? transformDiagnostics) {
        var children = originalSyntaxTree?.Children ?? Array.Empty<MarkdownSyntaxNode>();
        var topLevelBlocks = document.TopLevelBlocks;
        var spans = new List<MarkdownSourceSpan?>(document.Blocks.Count);

        if (children.Count > 0) {
            var childCount = Math.Min(children.Count, topLevelBlocks.Count);
            for (int i = 0; i < childCount; i++) {
                if (topLevelBlocks[i] is FrontMatterBlock) {
                    continue;
                }

                spans.Add(children[i].SourceSpan);
            }
        }

        while (spans.Count < document.Blocks.Count) {
            spans.Add(null);
        }

        if (transformDiagnostics != null) {
            for (int i = 0; i < transformDiagnostics.Count; i++) {
                spans = UpdateTopLevelBlockSourceSpans(spans, transformDiagnostics[i]);
            }
        }

        if (spans.Count > document.Blocks.Count) {
            spans.RemoveRange(document.Blocks.Count, spans.Count - document.Blocks.Count);
        }

        return spans;
    }

    private static MarkdownSourceSpan? GetFrontMatterSpan(MarkdownDoc document, MarkdownSyntaxNode originalSyntaxTree) {
        if (document?.DocumentHeader == null || originalSyntaxTree == null || originalSyntaxTree.Children.Count == 0) {
            return null;
        }

        var first = originalSyntaxTree.Children[0];
        return first.Kind == MarkdownSyntaxKind.FrontMatter ? first.SourceSpan : null;
    }

    private static List<MarkdownSourceSpan?> UpdateTopLevelBlockSourceSpans(
        IReadOnlyList<MarkdownSourceSpan?> previous,
        MarkdownDocumentTransformDiagnostic diagnostic) {
        var updated = new List<MarkdownSourceSpan?>(diagnostic.BlockCountAfter);
        var prefixCount = Math.Min(diagnostic.ChangedBlockStartBefore, previous.Count);
        for (var i = 0; i < prefixCount; i++) {
            updated.Add(previous[i]);
        }

        for (var i = 0; i < diagnostic.ChangedBlockCountAfter; i++) {
            updated.Add(diagnostic.AffectedSourceSpan);
        }

        var suffixCount = previous.Count - diagnostic.ChangedBlockStartBefore - diagnostic.ChangedBlockCountBefore;
        var suffixStart = Math.Max(prefixCount, previous.Count - suffixCount);
        for (var i = suffixStart; i < previous.Count; i++) {
            updated.Add(previous[i]);
        }

        while (updated.Count < diagnostic.BlockCountAfter) {
            updated.Add(null);
        }

        if (updated.Count > diagnostic.BlockCountAfter) {
            updated.RemoveRange(diagnostic.BlockCountAfter, updated.Count - diagnostic.BlockCountAfter);
        }

        return updated;
    }
}
