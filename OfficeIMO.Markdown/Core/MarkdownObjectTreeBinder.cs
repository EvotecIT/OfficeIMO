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
    }

    internal static IReadOnlyList<MarkdownObject> GetChildObjects(MarkdownObject parent) {
        if (parent == null) {
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
                var listItemBlocks = listItem.BlockChildren;
                for (int i = 0; i < listItemBlocks.Count; i++) {
                    if (listItemBlocks[i] is MarkdownObject listItemChild) {
                        yield return listItemChild;
                    }
                }
                yield break;

            case TableBlock table:
                var headerCells = table.HeaderCells;
                for (int i = 0; i < headerCells.Count; i++) {
                    yield return headerCells[i];
                }

                var rows = table.RowCells;
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    var row = rows[rowIndex];
                    for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                        yield return row[columnIndex];
                    }
                }
                yield break;

            case TableCell cell:
                for (int i = 0; i < cell.Blocks.Count; i++) {
                    if (cell.Blocks[i] is MarkdownObject cellBlock) {
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
                for (int i = 0; i < definitionGroup.Terms.Count; i++) {
                    yield return definitionGroup.Terms[i];
                }

                for (int i = 0; i < definitionGroup.Definitions.Count; i++) {
                    yield return definitionGroup.Definitions[i];
                }
                yield break;

            case DefinitionListEntry definitionEntry:
                yield return definitionEntry.Term;
                yield return definitionEntry.Definition;
                yield break;

            case DefinitionListDefinition definition:
                for (int i = 0; i < definition.Blocks.Count; i++) {
                    if (definition.Blocks[i] is MarkdownObject definitionBlock) {
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

    private static void MapSourceSpans(MarkdownSyntaxNode syntaxNode) {
        if (syntaxNode.AssociatedObject is MarkdownObject markdownObject) {
            markdownObject.SourceSpan = syntaxNode.SourceSpan;
        }

        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            MapSourceSpans(syntaxNode.Children[i]);
        }
    }
}
