namespace OfficeIMO.Markdown;

internal static class MarkdownObjectTreeBinder {
    internal static void BindDocument(MarkdownDocument document, MarkdownSyntaxNode? syntaxTree = null) {
        ArgumentNullException.ThrowIfNull(document);

        BindObject(document, parent: null, indexInParent: null, previousSibling: null, nextSibling: null);

        if (syntaxTree != null) {
            MapSourceSpans(syntaxTree);
        }

        document.MarkObjectTreeBound();
    }

    internal static void BindSourceSpans(MarkdownSyntaxNode syntaxNode) {
        ArgumentNullException.ThrowIfNull(syntaxNode);

        MapSourceSpans(syntaxNode);
    }

    internal static IReadOnlyList<MarkdownObject> GetChildObjects(MarkdownObject parent) {
        if (parent == null || !CanContainChildObjects(parent)) {
            return Array.Empty<MarkdownObject>();
        }

        var children = new List<MarkdownObject>();
        foreach (var child in EnumerateChildObjects(parent)) {
            if (child != null) {
                children.Add(child);
            }
        }

        return children;
    }

    private static bool CanContainChildObjects(MarkdownObject parent) =>
        parent is MarkdownDocument ||
        parent is HeadingBlock ||
        parent is CalloutBlock ||
        parent is DetailsBlock ||
        parent is IMarkdownListBlock ||
        parent is ListItem ||
        parent is TableBlock ||
        parent is TableRow ||
        parent is TableCell ||
        parent is DefinitionListBlock ||
        parent is DefinitionListGroup ||
        parent is DefinitionListTerm ||
        parent is DefinitionListEntry ||
        parent is DefinitionListDefinition ||
        parent is InlineSequence ||
        parent is IInlineSyntaxMarkdownBlock ||
        parent is IInlineContainerMarkdownInline ||
        parent is IChildMarkdownBlockContainer;

    internal static IEnumerable<MarkdownObject> EnumerateChildObjects(MarkdownObject parent) {
        switch (parent) {
            case MarkdownDocument document:
                if (document.DocumentHeader is MarkdownObject headerObject) {
                    yield return headerObject;
                }

                for (int i = 0; i < document.Blocks.Count; i++) {
                    if (document.Blocks[i] is MarkdownObject blockObject) {
                        yield return blockObject;
                    }
                }
                yield break;

            case HeadingBlock heading:
                yield return heading.Inlines;
                yield break;

            case CalloutBlock callout:
                yield return callout.TitleInlines;
                for (int i = 0; i < callout.ChildBlocks.Count; i++) {
                    if (callout.ChildBlocks[i] is MarkdownObject calloutChild) {
                        yield return calloutChild;
                    }
                }
                yield break;

            case DetailsBlock details:
                if (details.Summary is MarkdownObject summaryObject) {
                    yield return summaryObject;
                }

                for (int i = 0; i < details.ChildBlocks.Count; i++) {
                    if (details.ChildBlocks[i] is MarkdownObject detailsChild) {
                        yield return detailsChild;
                    }
                }
                yield break;

            case IMarkdownListBlock listBlock:
                for (int i = 0; i < listBlock.ListItems.Count; i++) {
                    yield return listBlock.ListItems[i];
                }
                yield break;

            case ListItem listItem:
                var listItemBlocks = listItem.ChildBlocks;
                for (int i = 0; i < listItemBlocks.Count; i++) {
                    if (listItemBlocks[i] is MarkdownObject listItemChild) {
                        yield return listItemChild;
                    }
                }
                yield break;

            case TableBlock table:
                var headerRow = table.HeaderRow;
                if (headerRow != null) {
                    yield return headerRow;
                }

                var rows = table.BodyRows;
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    yield return rows[rowIndex];
                }
                yield break;

            case TableRow row:
                for (int i = 0; i < row.Cells.Count; i++) {
                    yield return row.Cells[i];
                }
                yield break;

            case TableCell cell:
                for (int i = 0; i < cell.ChildBlocks.Count; i++) {
                    if (cell.ChildBlocks[i] is MarkdownObject cellBlock) {
                        yield return cellBlock;
                    }
                }
                yield break;

            case DefinitionListBlock definitionList:
                for (int i = 0; i < definitionList.Groups.Count; i++) {
                    yield return definitionList.Groups[i];
                }
                yield break;

            case DefinitionListGroup definitionGroup:
                for (int i = 0; i < definitionGroup.TermItems.Count; i++) {
                    yield return definitionGroup.TermItems[i];
                }

                for (int i = 0; i < definitionGroup.Definitions.Count; i++) {
                    yield return definitionGroup.Definitions[i];
                }
                yield break;

            case DefinitionListTerm term:
                yield return term.Inlines;
                yield break;

            case DefinitionListEntry definitionEntry:
                yield return definitionEntry.Term;
                yield return definitionEntry.Definition;
                yield break;

            case DefinitionListDefinition definition:
                for (int i = 0; i < definition.ChildBlocks.Count; i++) {
                    if (definition.ChildBlocks[i] is MarkdownObject definitionBlock) {
                        yield return definitionBlock;
                    }
                }
                yield break;

            case InlineSequence inlineSequence:
                for (int i = 0; i < inlineSequence.Nodes.Count; i++) {
                    if (inlineSequence.Nodes[i] is MarkdownObject inlineNode) {
                        yield return inlineNode;
                    }
                }
                yield break;
        }

        if (parent is IInlineSyntaxMarkdownBlock inlineBlock && inlineBlock.SyntaxInlines is MarkdownObject inlineBlockObject) {
            yield return inlineBlockObject;
        }

        if (parent is IInlineContainerMarkdownInline inlineContainer && inlineContainer.NestedInlines is MarkdownObject nestedInlines) {
            yield return nestedInlines;
        }

        if (parent is IChildMarkdownBlockContainer container) {
            for (int i = 0; i < container.ChildBlocks.Count; i++) {
                if (container.ChildBlocks[i] is MarkdownObject childBlock) {
                    yield return childBlock;
                }
            }
        }
    }

    // Parsed and hand-built trees can be deeper than the call stack allows, so both passes walk
    // the tree with an explicit stack. Children are pushed in reverse to keep pre-order binding.
    private readonly struct BindWork {
        internal BindWork(MarkdownObject node, MarkdownObject? parent, int? indexInParent, MarkdownObject? previousSibling, MarkdownObject? nextSibling) {
            Node = node;
            Parent = parent;
            IndexInParent = indexInParent;
            PreviousSibling = previousSibling;
            NextSibling = nextSibling;
        }

        internal MarkdownObject Node { get; }
        internal MarkdownObject? Parent { get; }
        internal int? IndexInParent { get; }
        internal MarkdownObject? PreviousSibling { get; }
        internal MarkdownObject? NextSibling { get; }
    }

    private static void BindObject(
        MarkdownObject node,
        MarkdownObject? parent,
        int? indexInParent,
        MarkdownObject? previousSibling,
        MarkdownObject? nextSibling) {
        var pending = new Stack<BindWork>();
        var children = new List<BindWork>();
        pending.Push(new BindWork(node, parent, indexInParent, previousSibling, nextSibling));

        while (pending.Count > 0) {
            BindWork work = pending.Pop();
            work.Node.SetTreePosition(work.Parent, work.IndexInParent, work.PreviousSibling, work.NextSibling);

            children.Clear();
            if (!TryCollectKnownChildren(work.Node, children)) {
                var childObjects = GetChildObjects(work.Node);
                for (int i = 0; i < childObjects.Count; i++) {
                    children.Add(new BindWork(
                        childObjects[i],
                        work.Node,
                        i,
                        i > 0 ? childObjects[i - 1] : null,
                        i + 1 < childObjects.Count ? childObjects[i + 1] : null));
                }
            }

            for (int i = children.Count - 1; i >= 0; i--) {
                pending.Push(children[i]);
            }
        }
    }

    private static bool TryCollectKnownChildren(MarkdownObject parent, List<BindWork> output) {
        switch (parent) {
            case MarkdownDocument document when document.DocumentHeader == null:
                CollectChildren(parent, leadingChild: null, document.Blocks, output);
                return true;

            case HeadingBlock heading:
                CollectOnlyChild(parent, heading.Inlines, output);
                return true;

            case IMarkdownListBlock listBlock:
                CollectChildren(parent, leadingChild: null, listBlock.ListItems, output);
                return true;

            case ListItem listItem when !listItem.HasAdditionalParagraphs:
                CollectChildren(
                    parent,
                    listItem.Content.Nodes.Count > 0 || listItem.NestedBlockCount == 0
                        ? listItem.LeadParagraphBlock
                        : null,
                    listItem.NestedBlocksOrEmpty,
                    output);
                return true;

            case ListItem listItem:
                CollectChildren(parent, leadingChild: null, listItem.ChildBlocks, output);
                return true;

            case TableBlock table:
                CollectChildren(parent, table.HeaderRow, table.BodyRows, output);
                return true;

            case TableRow row:
                CollectChildren(parent, leadingChild: null, row.Cells, output);
                return true;

            case TableCell cell:
                CollectChildren(parent, leadingChild: null, cell.ChildBlocks, output);
                return true;

            case InlineSequence inlineSequence:
                CollectChildren(parent, leadingChild: null, inlineSequence.Nodes, output);
                return true;

            case IInlineSyntaxMarkdownBlock inlineBlock:
                CollectOnlyChild(parent, inlineBlock.SyntaxInlines, output);
                return true;

            case IInlineContainerMarkdownInline inlineContainer
                when inlineContainer is not IChildMarkdownBlockContainer
                && inlineContainer is not ISyntaxChildrenMarkdownBlock:
                if (inlineContainer.NestedInlines is MarkdownObject nestedInlines) {
                    CollectOnlyChild(parent, nestedInlines, output);
                }
                return true;

            default:
                return false;
        }
    }

    private static void CollectOnlyChild(MarkdownObject parent, MarkdownObject child, List<BindWork> output) {
        output.Add(new BindWork(child, parent, indexInParent: 0, previousSibling: null, nextSibling: null));
    }

    private static void CollectChildren<T>(MarkdownObject parent, MarkdownObject? leadingChild, IReadOnlyList<T> children, List<BindWork> output) {
        MarkdownObject? previousSibling = null;
        MarkdownObject? pendingChild = leadingChild;
        int childIndex = leadingChild == null ? 0 : 1;

        for (int i = 0; i < children.Count; i++) {
            if (children[i] is not MarkdownObject child) {
                continue;
            }

            if (pendingChild != null) {
                output.Add(new BindWork(pendingChild, parent, childIndex - 1, previousSibling, child));
                previousSibling = pendingChild;
            }

            pendingChild = child;
            childIndex++;
        }

        if (pendingChild != null) {
            output.Add(new BindWork(pendingChild, parent, childIndex - 1, previousSibling, nextSibling: null));
        }
    }

    private static void MapSourceSpans(MarkdownSyntaxNode root) {
        var pending = new Stack<MarkdownSyntaxNode>();
        pending.Push(root);

        while (pending.Count > 0) {
            MarkdownSyntaxNode syntaxNode = pending.Pop();
            if (syntaxNode.AssociatedObject is MarkdownObject markdownObject) {
                markdownObject.BindSyntaxNode(syntaxNode);
                if (markdownObject is MarkdownInline inline) {
                    MarkdownInlineMetadataSourceSpans.ReleaseRedundantFormattingMetadata(inline);
                }
            }

            for (int i = syntaxNode.Children.Count - 1; i >= 0; i--) {
                pending.Push(syntaxNode.Children[i]);
            }
        }
    }
}
