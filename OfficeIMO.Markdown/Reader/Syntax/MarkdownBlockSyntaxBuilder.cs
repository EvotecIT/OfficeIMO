namespace OfficeIMO.Markdown;

internal static class MarkdownBlockSyntaxBuilder {
    private static readonly MarkdownBlockSyntaxBuilderContext _context = new();

    internal static MarkdownSyntaxNode BuildBlock(IMarkdownBlock block, MarkdownSourceSpan? span = null, bool isGenerated = false) {
        var effectiveSpan = span ?? (block as MarkdownObject)?.SourceSpan;
        MarkdownSyntaxNode node;

        if (block is ISyntaxMarkdownBlockWithContext syntaxBlockWithContext) {
            node = ApplyBlockAttributes(block, syntaxBlockWithContext.BuildSyntaxNode(_context, effectiveSpan));
        } else if (block is ISyntaxMarkdownBlock syntaxBlock) {
            node = ApplyBlockAttributes(block, syntaxBlock.BuildSyntaxNode(effectiveSpan));
        } else {
            node = new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Unknown,
                effectiveSpan,
                block.RenderMarkdown(),
                associatedObject: block,
                attributes: (block as MarkdownObject)?.Attributes);
        }

        return isGenerated && !node.IsGenerated
            ? CloneSyntaxNodePreservingSourceSpan(node, isGenerated: true)
            : node;
    }

    internal static MarkdownSyntaxNode BuildInlineBlock(IInlineSyntaxMarkdownBlock inlineBlock, MarkdownSourceSpan? span = null, bool isGenerated = false) {
        var children = MarkdownInlineSyntaxBuilder.BuildChildren(inlineBlock.SyntaxInlines);
        return new MarkdownSyntaxNode(
            inlineBlock.SyntaxKind,
            span ?? inlineBlock.ProvidedSyntaxSpan ?? GetAggregateSpan(children),
            inlineBlock.SyntaxInlines.RenderMarkdown(),
            children: children,
            associatedObject: inlineBlock,
            isGenerated: isGenerated);
    }

    internal static MarkdownSyntaxNode BuildInlineContainerNode(
        MarkdownSyntaxKind kind,
        InlineSequence inlines,
        MarkdownSourceSpan? span = null,
        string? literal = null,
        object? associatedObject = null,
        bool isGenerated = false) {
        var markdownObject = associatedObject as MarkdownObject;
        var inlineChildren = MarkdownInlineSyntaxBuilder.BuildChildren(inlines);
        var children = MarkdownGenericAttributeSyntaxNodes.Append(inlineChildren, markdownObject);
        return new MarkdownSyntaxNode(
            kind,
            MarkdownGenericAttributeSyntaxNodes.GetContainingSpan(span, children) ?? GetAggregateSpan(children),
            literal ?? inlines?.RenderMarkdown(),
            children: children,
            associatedObject: associatedObject ?? inlines,
            attributes: markdownObject?.Attributes,
            isGenerated: isGenerated);
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> BuildChildSyntaxNodes(IEnumerable<IMarkdownBlock> children) {
        var nodes = new List<MarkdownSyntaxNode>();
        foreach (var child in children) {
            if (child == null) continue;
            nodes.Add(BuildBlock(child));
        }
        return nodes;
    }

    internal static bool ChildSyntaxNodesMatchBlocks(
        IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren,
        IReadOnlyList<IMarkdownBlock> blocks) {
        if (syntaxChildren == null || blocks == null || syntaxChildren.Count != blocks.Count) {
            return false;
        }

        for (int i = 0; i < blocks.Count; i++) {
            if (syntaxChildren[i] == null || !ReferenceEquals(syntaxChildren[i].AssociatedObject, blocks[i])) {
                return false;
            }
        }

        return true;
    }

    internal static IReadOnlyList<MarkdownSyntaxNode> BuildCanonicalChildSyntaxNodes(
        IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren,
        IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var children = new List<MarkdownSyntaxNode>(blocks.Count);
        for (int i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            var cachedSyntax = FindCanonicalSyntaxChild(syntaxChildren, block, i);
            if (cachedSyntax != null &&
                IsSyntaxChildForBlock(cachedSyntax, block)) {
                children.Add(CloneSyntaxNode(cachedSyntax));
                continue;
            }

            if (cachedSyntax != null &&
                CanBuildSourceBackedCompatibleSyntax(cachedSyntax, block)) {
                children.Add(BuildBlock(block, cachedSyntax.SourceSpan, isGenerated: false));
                continue;
            }

            children.Add(BuildBlock(block, cachedSyntax?.SourceSpan, isGenerated: true));
        }

        return children;
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

    private static MarkdownSyntaxNode? FindCanonicalSyntaxChild(
        IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren,
        IMarkdownBlock block,
        int preferredIndex) {
        if (syntaxChildren == null || syntaxChildren.Count == 0) {
            return null;
        }

        var preferredSyntax = GetSyntaxChildOrNull(syntaxChildren, preferredIndex);
        if (IsSyntaxChildForBlock(preferredSyntax, block)) {
            return preferredSyntax;
        }

        for (int i = 0; i < syntaxChildren.Count; i++) {
            if (IsSyntaxChildForBlock(syntaxChildren[i], block)) {
                return syntaxChildren[i];
            }
        }

        if (IsSyntaxChildCompatibleWithBlock(preferredSyntax, block)) {
            return preferredSyntax;
        }

        for (int i = 0; i < syntaxChildren.Count; i++) {
            if (IsSyntaxChildCompatibleWithBlock(syntaxChildren[i], block)) {
                return syntaxChildren[i];
            }
        }

        return null;
    }

    private static MarkdownSyntaxNode? GetSyntaxChildOrNull(IReadOnlyList<MarkdownSyntaxNode> syntaxChildren, int index) =>
        index >= 0 && index < syntaxChildren.Count ? syntaxChildren[index] : null;

    private static bool IsSyntaxChildForBlock(MarkdownSyntaxNode? syntaxNode, IMarkdownBlock block) =>
        syntaxNode?.AssociatedObject != null && ReferenceEquals(syntaxNode.AssociatedObject, block);

    private static bool CanBuildSourceBackedCompatibleSyntax(MarkdownSyntaxNode syntaxNode, IMarkdownBlock block) {
        if (syntaxNode.IsGenerated ||
            !syntaxNode.SourceSpan.HasValue ||
            !IsSyntaxChildCompatibleWithBlock(syntaxNode, block) ||
            syntaxNode.Literal == null) {
            return false;
        }

        return string.Equals(
            NormalizeSyntaxLiteralLineEndings(syntaxNode.Literal),
            NormalizeSyntaxLiteralLineEndings(block.RenderMarkdown()),
            StringComparison.Ordinal);
    }

    internal static MarkdownSyntaxNode CloneSyntaxNode(MarkdownSyntaxNode node) {
        return CloneSyntaxNode(node, isGeneratedOverride: null);
    }

    private static MarkdownSyntaxNode CloneSyntaxNodePreservingSourceSpan(MarkdownSyntaxNode node, bool isGenerated) {
        var children = node.Children.Count == 0
            ? Array.Empty<MarkdownSyntaxNode>()
            : node.Children.Select(CloneSyntaxNode).ToArray();

        return new MarkdownSyntaxNode(
            node.Kind,
            node.SourceSpan,
            node.Literal,
            children,
            node.AssociatedObject,
            node.CustomKind,
            node.Attributes,
            isGenerated);
    }

    private static MarkdownSyntaxNode CloneSyntaxNode(MarkdownSyntaxNode node, bool? isGeneratedOverride) {
        var children = node.Children.Count == 0
            ? Array.Empty<MarkdownSyntaxNode>()
            : node.Children.Select(CloneSyntaxNode).ToArray();

        return new MarkdownSyntaxNode(
            node.Kind,
            MarkdownGenericAttributeSyntaxNodes.GetContainingSpan(node.SourceSpan, children),
            node.Literal,
            children,
            node.AssociatedObject,
            node.CustomKind,
            node.Attributes,
            isGeneratedOverride ?? node.IsGenerated);
    }

    private static MarkdownSyntaxNode ApplyBlockAttributes(IMarkdownBlock block, MarkdownSyntaxNode node) {
        if (node == null) {
            throw new InvalidOperationException("Markdown syntax block builders must return a syntax node.");
        }

        if (block is not MarkdownObject markdownObject ||
            markdownObject.Attributes.IsEmpty) {
            return node;
        }

        var clonedChildren = node.Children.Count == 0
            ? Array.Empty<MarkdownSyntaxNode>()
            : node.Children.Select(CloneSyntaxNode).ToArray();
        var children = MarkdownGenericAttributeSyntaxNodes.Append(clonedChildren, markdownObject);

        return new MarkdownSyntaxNode(
            node.Kind,
            MarkdownGenericAttributeSyntaxNodes.GetContainingSpan(node.SourceSpan, children) ?? node.SourceSpan,
            node.Literal,
            children,
            node.AssociatedObject,
            node.CustomKind,
            node.Attributes.IsEmpty ? markdownObject.Attributes : node.Attributes,
            node.IsGenerated);
    }

    private static bool IsSyntaxChildCompatibleWithBlock(MarkdownSyntaxNode? syntaxNode, IMarkdownBlock block) {
        if (syntaxNode == null) {
            return false;
        }

        return block switch {
            ParagraphBlock => syntaxNode.Kind == MarkdownSyntaxKind.Paragraph,
            HeadingBlock => syntaxNode.Kind == MarkdownSyntaxKind.Heading,
            QuoteBlock => syntaxNode.Kind == MarkdownSyntaxKind.Quote,
            UnorderedListBlock => syntaxNode.Kind == MarkdownSyntaxKind.UnorderedList,
            OrderedListBlock => syntaxNode.Kind == MarkdownSyntaxKind.OrderedList,
            CodeBlock => syntaxNode.Kind == MarkdownSyntaxKind.CodeBlock,
            SemanticFencedBlock => syntaxNode.Kind == MarkdownSyntaxKind.SemanticFencedBlock,
            TableBlock => syntaxNode.Kind == MarkdownSyntaxKind.Table,
            HorizontalRuleBlock => syntaxNode.Kind == MarkdownSyntaxKind.HorizontalRule,
            ImageBlock => syntaxNode.Kind == MarkdownSyntaxKind.Image,
            CalloutBlock => syntaxNode.Kind == MarkdownSyntaxKind.Callout,
            DefinitionListBlock => syntaxNode.Kind == MarkdownSyntaxKind.DefinitionList,
            FootnoteDefinitionBlock => syntaxNode.Kind == MarkdownSyntaxKind.FootnoteDefinition,
            DetailsBlock => syntaxNode.Kind == MarkdownSyntaxKind.Details,
            SummaryBlock => syntaxNode.Kind == MarkdownSyntaxKind.Summary,
            FrontMatterBlock => syntaxNode.Kind == MarkdownSyntaxKind.FrontMatter,
            HtmlRawBlock => syntaxNode.Kind == MarkdownSyntaxKind.HtmlRaw,
            HtmlCommentBlock => syntaxNode.Kind == MarkdownSyntaxKind.HtmlComment,
            TocBlock => syntaxNode.Kind == MarkdownSyntaxKind.Toc,
            TocMarkerBlock => syntaxNode.Kind == MarkdownSyntaxKind.Toc,
            _ => false
        };
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
