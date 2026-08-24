namespace OfficeIMO.Markdown;

internal static class MarkdownObjectTreeBinder {
    internal static void BindDocument(MarkdownDoc document, MarkdownSyntaxNode? syntaxTree = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        BindObject(document, parent: null, indexInParent: null, previousSibling: null, nextSibling: null);

        if (syntaxTree != null) {
            MapSourceSpans(syntaxTree);
        }

        document.MarkObjectTreeBound();
    }

    internal static void BindSourceSpans(MarkdownSyntaxNode syntaxNode) {
        if (syntaxNode == null) {
            throw new ArgumentNullException(nameof(syntaxNode));
        }

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
        parent is MarkdownDoc ||
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
            case MarkdownDoc document:
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

    private static void BindObject(
        MarkdownObject node,
        MarkdownObject? parent,
        int? indexInParent,
        MarkdownObject? previousSibling,
        MarkdownObject? nextSibling) {
        node.SetTreePosition(parent, indexInParent, previousSibling, nextSibling);

        if (TryBindKnownChildren(node)) {
            return;
        }

        var children = GetChildObjects(node);
        for (int i = 0; i < children.Count; i++) {
            BindObject(
                children[i],
                node,
                i,
                i > 0 ? children[i - 1] : null,
                i + 1 < children.Count ? children[i + 1] : null);
        }
    }

    private static bool TryBindKnownChildren(MarkdownObject parent) {
        switch (parent) {
            case MarkdownDoc document when document.DocumentHeader == null:
                BindChildren(parent, document.Blocks);
                return true;

            case HeadingBlock heading:
                BindOnlyChild(parent, heading.Inlines);
                return true;

            case IMarkdownListBlock listBlock:
                BindChildren(parent, listBlock.ListItems);
                return true;

            case ListItem listItem when !listItem.HasAdditionalParagraphs:
                BindChildren(
                    parent,
                    listItem.Content.Nodes.Count > 0 || listItem.NestedBlockCount == 0
                        ? listItem.LeadParagraphBlock
                        : null,
                    listItem.NestedBlocksOrEmpty);
                return true;

            case ListItem listItem:
                BindChildren(parent, listItem.ChildBlocks);
                return true;

            case TableBlock table:
                BindChildren(parent, table.HeaderRow, table.BodyRows);
                return true;

            case TableRow row:
                BindChildren(parent, row.Cells);
                return true;

            case TableCell cell:
                BindChildren(parent, cell.ChildBlocks);
                return true;

            case InlineSequence inlineSequence:
                BindChildren(parent, inlineSequence.Nodes);
                return true;

            case IInlineSyntaxMarkdownBlock inlineBlock:
                BindOnlyChild(parent, inlineBlock.SyntaxInlines);
                return true;

            case IInlineContainerMarkdownInline inlineContainer
                when inlineContainer is not IChildMarkdownBlockContainer
                && inlineContainer is not ISyntaxChildrenMarkdownBlock:
                if (inlineContainer.NestedInlines is MarkdownObject nestedInlines) {
                    BindOnlyChild(parent, nestedInlines);
                }
                return true;

            default:
                return false;
        }
    }

    private static void BindOnlyChild(MarkdownObject parent, MarkdownObject child) {
        BindObject(child, parent, indexInParent: 0, previousSibling: null, nextSibling: null);
    }

    private static void BindChildren<T>(MarkdownObject parent, MarkdownObject? leadingChild, IReadOnlyList<T> children) {
        MarkdownObject? previousSibling = null;
        MarkdownObject? pendingChild = leadingChild;
        int childIndex = leadingChild == null ? 0 : 1;

        for (int i = 0; i < children.Count; i++) {
            if (children[i] is not MarkdownObject child) {
                continue;
            }

            if (pendingChild != null) {
                BindObject(pendingChild, parent, childIndex - 1, previousSibling, child);
                previousSibling = pendingChild;
            }

            pendingChild = child;
            childIndex++;
        }

        if (pendingChild != null) {
            BindObject(pendingChild, parent, childIndex - 1, previousSibling, nextSibling: null);
        }
    }

    private static void BindChildren<T>(MarkdownObject parent, IReadOnlyList<T> children) {
        MarkdownObject? previousSibling = null;
        MarkdownObject? pendingChild = null;
        int childIndex = 0;

        for (int i = 0; i < children.Count; i++) {
            if (children[i] is not MarkdownObject child) {
                continue;
            }

            if (pendingChild != null) {
                BindObject(pendingChild, parent, childIndex - 1, previousSibling, child);
                previousSibling = pendingChild;
            }

            pendingChild = child;
            childIndex++;
        }

        if (pendingChild != null) {
            BindObject(pendingChild, parent, childIndex - 1, previousSibling, nextSibling: null);
        }
    }

    private static void MapSourceSpans(MarkdownSyntaxNode syntaxNode) {
        if (syntaxNode.AssociatedObject is MarkdownObject markdownObject) {
            markdownObject.BindSyntaxNode(syntaxNode);
        }

        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            MapSourceSpans(syntaxNode.Children[i]);
        }
    }
}
