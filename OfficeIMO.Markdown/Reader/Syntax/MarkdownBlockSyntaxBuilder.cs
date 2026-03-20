namespace OfficeIMO.Markdown;

internal static class MarkdownBlockSyntaxBuilder {
    internal static MarkdownSyntaxNode BuildBlock(IMarkdownBlock block, MarkdownSourceSpan? span = null) {
        if (block is ISyntaxMarkdownBlock syntaxBlock) {
            return syntaxBlock.BuildSyntaxNode(span);
        }

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.Unknown, span, block.RenderMarkdown(), associatedObject: block);
    }

    internal static MarkdownSyntaxNode BuildInlineBlock(IInlineSyntaxMarkdownBlock inlineBlock, MarkdownSourceSpan? span = null) {
        var children = MarkdownInlineSyntaxBuilder.BuildChildren(inlineBlock.SyntaxInlines);
        return new MarkdownSyntaxNode(
            inlineBlock.SyntaxKind,
            span ?? inlineBlock.ProvidedSyntaxSpan ?? GetAggregateSpan(children),
            inlineBlock.SyntaxInlines.RenderMarkdown(),
            children: children,
            associatedObject: inlineBlock);
    }

    internal static MarkdownSyntaxNode BuildInlineContainerNode(
        MarkdownSyntaxKind kind,
        InlineSequence inlines,
        MarkdownSourceSpan? span = null,
        string? literal = null) {
        var children = MarkdownInlineSyntaxBuilder.BuildChildren(inlines);
        return new MarkdownSyntaxNode(
            kind,
            span ?? GetAggregateSpan(children),
            literal ?? inlines?.RenderMarkdown(),
            children: children,
            associatedObject: inlines);
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> BuildChildSyntaxNodes(IEnumerable<IMarkdownBlock> children) {
        var nodes = new List<MarkdownSyntaxNode>();
        foreach (var child in children) {
            if (child == null) continue;
            nodes.Add(BuildBlock(child));
        }
        return nodes;
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> GetOwnedSyntaxChildrenOrBuild(IChildMarkdownBlockContainer block) {
        if (block is IOwnedSyntaxChildrenMarkdownBlock ownedSyntaxChildren) {
            return ownedSyntaxChildren.BuildOwnedSyntaxChildren();
        }

        if (block is ISyntaxChildrenMarkdownBlock syntaxOwner &&
            syntaxOwner.ProvidedSyntaxChildren != null &&
            syntaxOwner.ProvidedSyntaxChildren.Count > 0) {
            return syntaxOwner.ProvidedSyntaxChildren;
        }

        return BuildChildSyntaxNodes(block.ChildBlocks);
    }

    internal static string NormalizeSyntaxLiteralLineEndings(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        string normalized = value!;
        return normalized.Replace("\r\n", "\n").Replace('\r', '\n');
    }

    internal static MarkdownSourceSpan? GetAggregateSpan(IReadOnlyList<MarkdownSyntaxNode> nodes) {
        if (nodes == null || nodes.Count == 0) return null;

        MarkdownSourceSpan? aggregate = null;
        for (int i = 0; i < nodes.Count; i++) {
            var span = nodes[i].SourceSpan;
            if (!span.HasValue) continue;

            aggregate = !aggregate.HasValue
                ? span
                : MergeSpans(aggregate.Value, span.Value);
        }

        return aggregate;
    }

    private static MarkdownSourceSpan MergeSpans(MarkdownSourceSpan left, MarkdownSourceSpan right) {
        var start = ComparePositions(left.StartLine, left.StartColumn, right.StartLine, right.StartColumn) <= 0 ? left : right;
        var end = ComparePositions(left.EndLine, left.EndColumn, right.EndLine, right.EndColumn) >= 0 ? left : right;

        if (start.StartColumn.HasValue && end.EndColumn.HasValue) {
            return new MarkdownSourceSpan(
                start.StartLine,
                start.StartColumn.Value,
                end.EndLine,
                end.EndColumn.Value,
                start.StartOffset,
                end.EndOffset);
        }

        return new MarkdownSourceSpan(start.StartLine, end.EndLine);
    }

    private static int ComparePositions(int lineA, int? columnA, int lineB, int? columnB) {
        var lineCompare = lineA.CompareTo(lineB);
        if (lineCompare != 0) {
            return lineCompare;
        }

        return (columnA ?? 1).CompareTo(columnB ?? 1);
    }
}
