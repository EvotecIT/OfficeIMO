namespace OfficeIMO.Markdown;

internal static class MarkdownBlockSyntaxBuilder {
    internal static MarkdownSyntaxNode BuildBlock(IMarkdownBlock block, MarkdownSourceSpan? span = null) {
        if (block is ISyntaxMarkdownBlock syntaxBlock) {
            return syntaxBlock.BuildSyntaxNode(span);
        }

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.Unknown, span, block.RenderMarkdown());
    }

    internal static MarkdownSyntaxNode BuildInlineBlock(IInlineSyntaxMarkdownBlock inlineBlock, MarkdownSourceSpan? span = null) =>
        new MarkdownSyntaxNode(
            inlineBlock.SyntaxKind,
            span ?? inlineBlock.ProvidedSyntaxSpan,
            inlineBlock.SyntaxInlines.RenderMarkdown());

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

        int? start = null;
        int? end = null;
        for (int i = 0; i < nodes.Count; i++) {
            var span = nodes[i].SourceSpan;
            if (!span.HasValue) continue;

            if (!start.HasValue || span.Value.StartLine < start.Value) start = span.Value.StartLine;
            if (!end.HasValue || span.Value.EndLine > end.Value) end = span.Value.EndLine;
        }

        if (!start.HasValue || !end.HasValue) return null;
        return new MarkdownSourceSpan(start.Value, end.Value);
    }
}
