namespace OfficeIMO.Markdown;

internal static class MarkdownObjectTreeBinder {
    internal static void BindDocument(MarkdownDocument document, MarkdownSyntaxNode? syntaxTree = null) =>
        BindDocument(document, syntaxTree, visited: null);

    /// <summary>
    /// Binds the document and, when <paramref name="visited"/> is supplied, records every bound object in
    /// binding order. Two documents with the same structure record corresponding objects at the same index.
    /// </summary>
    internal static void BindDocument(MarkdownDocument document, MarkdownSyntaxNode? syntaxTree, MarkdownObjectSequence? visited) {
        ArgumentNullException.ThrowIfNull(document);

        BindObject(document, parent: null, indexInParent: null, previousSibling: null, nextSibling: null, visited);

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

    // Parsed and hand-built trees can be deeper than the call stack allows, so binding walks the
    // tree with an explicit stack. A parent places its children when it is visited and children
    // are pushed in reverse, so containers are still visited in pre-order. Leaf text, code and
    // break nodes, most of a parsed tree, are placed without being visited.
    private static void BindObject(
        MarkdownObject node,
        MarkdownObject? parent,
        int? indexInParent,
        MarkdownObject? previousSibling,
        MarkdownObject? nextSibling,
        MarkdownObjectSequence? visited) {
        node.SetTreePosition(parent, indexInParent, previousSibling, nextSibling);
        visited?.Add(node);
        // The walk's scratch collections are kept for the next binding on this thread; a binding
        // that starts while one is running takes its own.
        var pending = t_pending ?? new Stack<MarkdownObject>();
        var children = t_children ?? new List<MarkdownObject>();
        t_pending = null;
        t_children = null;
        int largestPending = int.MaxValue;
        try {
            largestPending = BindObjectTree(node, visited, pending, children);
        } finally {
            pending.Clear();
            children.Clear();
            if (children.Capacity <= MaxRetainedScratchCapacity && largestPending <= MaxRetainedScratchCapacity) {
                t_children = children;
                t_pending = pending;
            }
        }
    }

    /// <summary>
    /// Places a block just appended to a bound document, so building a document block by block does not
    /// re-bind every earlier block. Returns false when the document has to be bound in full instead.
    /// </summary>
    internal static bool TryBindAppendedBlock(MarkdownDocument document, IMarkdownBlock appended) {
        if (appended is not MarkdownObject block) {
            return true;
        }

        IReadOnlyList<IMarkdownBlock> blocks = document.Blocks;
        if (blocks.Count == 0 || !ReferenceEquals(blocks[blocks.Count - 1], appended)) {
            return false;
        }

        // The previous child is the nearest earlier block that is an object, or the front matter.
        MarkdownObject? previous = null;
        for (int i = blocks.Count - 2; i >= 0 && previous == null; i--) {
            previous = blocks[i] as MarkdownObject;
        }

        previous ??= document.DocumentHeader;
        int index = 0;
        if (previous != null) {
            if (!ReferenceEquals(previous.Parent, document) || previous.NextSibling != null || previous.IndexInParent is not int previousIndex) {
                return false;
            }

            index = previousIndex + 1;
            previous.SetTreePosition(document, previousIndex, previous.PreviousSibling, block);
        }

        BindObject(block, document, index, previous, nextSibling: null, visited: null);
        return true;
    }

    private const int MaxRetainedScratchCapacity = 4096;
    [ThreadStatic] private static Stack<MarkdownObject>? t_pending;
    [ThreadStatic] private static List<MarkdownObject>? t_children;

    // Returns the largest number of objects waiting at once, which bounds the stack's capacity.
    private static int BindObjectTree(
        MarkdownObject node,
        MarkdownObjectSequence? visited,
        Stack<MarkdownObject> pending,
        List<MarkdownObject> children) {
        int largestPending = 1;
        pending.Push(node);

        while (pending.Count > 0) {
            MarkdownObject current = pending.Pop();

            children.Clear();
            if (!TryCollectKnownChildren(current, children)) {
                var childObjects = GetChildObjects(current);
                for (int i = 0; i < childObjects.Count; i++) {
                    childObjects[i].SetTreePosition(
                        current,
                        i,
                        i > 0 ? childObjects[i - 1] : null,
                        i + 1 < childObjects.Count ? childObjects[i + 1] : null);
                    children.Add(childObjects[i]);
                }
            }

            visited?.AddRange(children);
            for (int i = children.Count - 1; i >= 0; i--) {
                MarkdownObject child = children[i];
                if (!IsLeafObject(child)) {
                    pending.Push(child);
                }
            }

            if (pending.Count > largestPending) {
                largestPending = pending.Count;
            }
        }

        return largestPending;
    }

    // Sealed types that never hold child objects.
    private static bool IsLeafObject(MarkdownObject node) =>
        node is MarkdownTextRun or CodeSpanInline or SoftBreakInline or HardBreakInline or CodeBlock
            or ImageInline or HtmlRawInline or HorizontalRuleBlock or HtmlRawBlock;

    private static bool TryCollectKnownChildren(MarkdownObject parent, List<MarkdownObject> output) {
        switch (parent) {
            // The two most common containers come first; neither matches an earlier case below.
            case InlineSequence inlineSequence:
                CollectChildren(parent, leadingChild: null, inlineSequence.Nodes, output);
                return true;

            case ParagraphBlock paragraph:
                CollectOnlyChild(parent, ((IInlineSyntaxMarkdownBlock)paragraph).SyntaxInlines, output);
                return true;

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

            // Implements only IChildMarkdownBlockContainer among the child-bearing interfaces.
            case QuoteBlock quote:
                CollectChildren(parent, leadingChild: null, ((IChildMarkdownBlockContainer)quote).ChildBlocks, output);
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

    private static void CollectOnlyChild(MarkdownObject parent, MarkdownObject child, List<MarkdownObject> output) {
        child.SetTreePosition(parent, indexInParent: 0, previousSibling: null, nextSibling: null);
        output.Add(child);
    }

    private static void CollectChildren(MarkdownObject parent, MarkdownObject? leadingChild, IReadOnlyList<IMarkdownInline> children, List<MarkdownObject> output) {
        // A non-generic copy per element type: shared generic code over reference types looks up
        // its type arguments at run time, which made this the binder's hottest path.
        MarkdownObject? previousSibling = null;
        MarkdownObject? pendingChild = leadingChild;
        int childIndex = leadingChild == null ? 0 : 1;
        int count = children.Count;
        // Inline sequences store their nodes in InlineNodeList; reading it directly avoids interface calls.
        var nodes = children as InlineNodeList;

        for (int i = 0; i < count; i++) {
            if ((nodes != null ? nodes[i] : children[i]) is not MarkdownObject child) {
                continue;
            }

            if (pendingChild != null) {
                pendingChild.SetTreePosition(parent, childIndex - 1, previousSibling, child);
                output.Add(pendingChild);
                previousSibling = pendingChild;
            }

            pendingChild = child;
            childIndex++;
        }

        if (pendingChild != null) {
            pendingChild.SetTreePosition(parent, childIndex - 1, previousSibling, nextSibling: null);
            output.Add(pendingChild);
        }
    }
    private static void CollectChildren(MarkdownObject parent, MarkdownObject? leadingChild, IReadOnlyList<IMarkdownBlock> children, List<MarkdownObject> output) {
        MarkdownObject? previousSibling = null;
        MarkdownObject? pendingChild = leadingChild;
        int childIndex = leadingChild == null ? 0 : 1;
        int count = children.Count;
        var list = children as List<IMarkdownBlock>;

        for (int i = 0; i < count; i++) {
            if ((list != null ? list[i] : children[i]) is not MarkdownObject child) {
                continue;
            }

            if (pendingChild != null) {
                pendingChild.SetTreePosition(parent, childIndex - 1, previousSibling, child);
                output.Add(pendingChild);
                previousSibling = pendingChild;
            }

            pendingChild = child;
            childIndex++;
        }

        if (pendingChild != null) {
            pendingChild.SetTreePosition(parent, childIndex - 1, previousSibling, nextSibling: null);
            output.Add(pendingChild);
        }
    }
    private static void CollectChildren(MarkdownObject parent, MarkdownObject? leadingChild, IReadOnlyList<ListItem> children, List<MarkdownObject> output) {
        MarkdownObject? previousSibling = null;
        MarkdownObject? pendingChild = leadingChild;
        int childIndex = leadingChild == null ? 0 : 1;
        int count = children.Count;
        var list = children as List<ListItem>;

        for (int i = 0; i < count; i++) {
            if ((list != null ? list[i] : children[i]) is not MarkdownObject child) {
                continue;
            }

            if (pendingChild != null) {
                pendingChild.SetTreePosition(parent, childIndex - 1, previousSibling, child);
                output.Add(pendingChild);
                previousSibling = pendingChild;
            }

            pendingChild = child;
            childIndex++;
        }

        if (pendingChild != null) {
            pendingChild.SetTreePosition(parent, childIndex - 1, previousSibling, nextSibling: null);
            output.Add(pendingChild);
        }
    }
    private static void CollectChildren<T>(MarkdownObject parent, MarkdownObject? leadingChild, IReadOnlyList<T> children, List<MarkdownObject> output) {
        MarkdownObject? previousSibling = null;
        MarkdownObject? pendingChild = leadingChild;
        int childIndex = leadingChild == null ? 0 : 1;

        for (int i = 0; i < children.Count; i++) {
            if (children[i] is not MarkdownObject child) {
                continue;
            }

            if (pendingChild != null) {
                pendingChild.SetTreePosition(parent, childIndex - 1, previousSibling, child);
                output.Add(pendingChild);
                previousSibling = pendingChild;
            }

            pendingChild = child;
            childIndex++;
        }

        if (pendingChild != null) {
            pendingChild.SetTreePosition(parent, childIndex - 1, previousSibling, nextSibling: null);
            output.Add(pendingChild);
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
