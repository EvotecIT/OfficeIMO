namespace OfficeIMO.Markdown;

internal static class MarkdownGenericAttributeSyntaxNodes {
    internal static MarkdownSyntaxNode? Create(MarkdownObject? markdownObject) {
        if (markdownObject == null || markdownObject.Attributes.IsEmpty) {
            return null;
        }

        var sourceSpan = MarkdownGenericAttributeSourceSpans.GetSourceSpan(markdownObject);
        if (!sourceSpan.HasValue) {
            return null;
        }

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.GenericAttributeBlock,
            sourceSpan,
            literal: MarkdownGenericAttributeSourceSpans.GetSourceText(markdownObject) ?? string.Empty);
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> Append(
        IReadOnlyList<MarkdownSyntaxNode>? children,
        MarkdownObject? markdownObject) {
        var attributeNode = Create(markdownObject);
        if (attributeNode == null) {
            return children ?? Array.Empty<MarkdownSyntaxNode>();
        }

        var sourceChildren = children ?? Array.Empty<MarkdownSyntaxNode>();
        var result = new List<MarkdownSyntaxNode>(sourceChildren.Count + 1);
        for (var i = 0; i < sourceChildren.Count; i++) {
            result.Add(sourceChildren[i]);
        }

        result.Add(attributeNode);
        return result;
    }

    internal static MarkdownSourceSpan? GetContainingSpan(
        MarkdownSourceSpan? ownerSpan,
        IReadOnlyList<MarkdownSyntaxNode> children) {
        var aggregateSpan = MarkdownBlockSyntaxBuilder.GetAggregateSpan(children);
        if (!ownerSpan.HasValue) {
            return aggregateSpan;
        }

        if (!aggregateSpan.HasValue || ownerSpan.Value.Contains(aggregateSpan.Value)) {
            return ownerSpan;
        }

        return MergeSpans(ownerSpan.Value, aggregateSpan.Value);
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
