namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static MarkdownSyntaxNode BuildDocumentSyntaxTree(IReadOnlyList<MarkdownSyntaxNode> children, MarkdownDoc? document = null) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.Document, MarkdownBlockSyntaxBuilder.GetAggregateSpan(children), children: children, associatedObject: document);

    private static MarkdownSyntaxNode DetachOriginalSyntaxAssociations(MarkdownSyntaxNode node) {
        if (node == null) {
            throw new ArgumentNullException(nameof(node));
        }

        if (node.Children.Count == 0) {
            return new MarkdownSyntaxNode(node.Kind, node.SourceSpan, node.Literal, customKind: node.CustomKind);
        }

        var children = new MarkdownSyntaxNode[node.Children.Count];
        for (int i = 0; i < node.Children.Count; i++) {
            children[i] = DetachOriginalSyntaxAssociations(node.Children[i]);
        }

        return new MarkdownSyntaxNode(node.Kind, node.SourceSpan, node.Literal, children, customKind: node.CustomKind);
    }

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
        var finalSyntaxTree = NormalizeFinalSyntaxTreeSpans(BuildSyntaxTree(document, blockSpans, frontMatterSpan));
        MarkdownTransformDiagnosticSyntaxHelper.PopulateFinalBlockAnchors(transformDiagnostics, finalSyntaxTree);
        return finalSyntaxTree;
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
        var blockChildren = children.Where(static child => child.AssociatedObject is IMarkdownBlock).ToList();
        var topLevelBlocks = document.TopLevelBlocks;
        var spans = new List<MarkdownSourceSpan?>(document.Blocks.Count);

        if (blockChildren.Count > 0) {
            var childCount = Math.Min(blockChildren.Count, topLevelBlocks.Count);
            for (int i = 0; i < childCount; i++) {
                if (topLevelBlocks[i] is FrontMatterBlock) {
                    continue;
                }

                spans.Add(blockChildren[i].SourceSpan);
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

    private static MarkdownSyntaxNode NormalizeFinalSyntaxTreeSpans(MarkdownSyntaxNode node) {
        if (node == null) {
            throw new ArgumentNullException(nameof(node));
        }

        IReadOnlyList<MarkdownSyntaxNode> children = node.Children;
        if (node.Children.Count > 0) {
            var normalizedChildren = new List<MarkdownSyntaxNode>(node.Children.Count);
            for (var i = 0; i < node.Children.Count; i++) {
                normalizedChildren.Add(NormalizeFinalSyntaxTreeChild(node.SourceSpan, node.Children[i]));
            }

            children = normalizedChildren;
        }

        return new MarkdownSyntaxNode(node.Kind, node.SourceSpan, node.Literal, children, node.AssociatedObject, node.CustomKind);
    }

    private static MarkdownSyntaxNode NormalizeFinalSyntaxTreeChild(MarkdownSourceSpan? parentSpan, MarkdownSyntaxNode child) {
        if (child == null) {
            throw new ArgumentNullException(nameof(child));
        }

        IReadOnlyList<MarkdownSyntaxNode> children = child.Children;
        if (child.Children.Count > 0) {
            var normalizedChildren = new List<MarkdownSyntaxNode>(child.Children.Count);
            for (var i = 0; i < child.Children.Count; i++) {
                normalizedChildren.Add(NormalizeFinalSyntaxTreeChild(child.SourceSpan, child.Children[i]));
            }

            children = normalizedChildren;
        }

        var normalizedSpan = child.SourceSpan;
        if (parentSpan.HasValue && normalizedSpan.HasValue && !parentSpan.Value.Contains(normalizedSpan.Value)) {
            var aggregateChildSpan = MarkdownBlockSyntaxBuilder.GetAggregateSpan(children);
            normalizedSpan = aggregateChildSpan.HasValue && parentSpan.Value.Contains(aggregateChildSpan.Value)
                ? aggregateChildSpan
                : null;
        }

        return new MarkdownSyntaxNode(child.Kind, normalizedSpan, child.Literal, children, child.AssociatedObject, child.CustomKind);
    }
}
