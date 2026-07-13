namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) MergeMarkdigDefinitionLazyListContinuations(
        IReadOnlyList<IMarkdownBlock> blocks,
        IReadOnlyList<MarkdownSyntaxNode> syntaxChildren,
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (blocks == null || syntaxChildren == null || blocks.Count < 2 || blocks.Count != syntaxChildren.Count) {
            return (blocks ?? Array.Empty<IMarkdownBlock>(), syntaxChildren ?? Array.Empty<MarkdownSyntaxNode>());
        }

        var mergedAny = false;
        var mergedBlocks = new List<IMarkdownBlock>(blocks.Count);
        var mergedSyntax = new List<MarkdownSyntaxNode>(syntaxChildren.Count);
        for (int index = 0; index < blocks.Count; index++) {
            var currentBlock = blocks[index];
            var currentSyntax = syntaxChildren[index];

            while (index + 1 < blocks.Count &&
                   TryMergeDefinitionLazyListContinuation(currentBlock, currentSyntax, blocks[index + 1], syntaxChildren[index + 1], sourceLines, options, state, out var combinedSyntax)) {
                currentSyntax = combinedSyntax;
                index++;
                mergedAny = true;
            }

            mergedBlocks.Add(currentBlock);
            mergedSyntax.Add(currentSyntax);
        }

        return mergedAny
            ? (mergedBlocks, mergedSyntax)
            : (blocks, syntaxChildren);
    }

    private static bool TryMergeDefinitionLazyListContinuation(
        IMarkdownBlock currentBlock,
        MarkdownSyntaxNode currentSyntax,
        IMarkdownBlock nextBlock,
        MarkdownSyntaxNode nextSyntax,
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out MarkdownSyntaxNode combinedSyntax) {
        combinedSyntax = currentSyntax;
        if (currentBlock is IMarkdownListBlock currentList &&
            nextBlock is ParagraphBlock nextParagraph &&
            nextSyntax.Kind == MarkdownSyntaxKind.Paragraph &&
            IsDefinitionLazyParagraphContinuationCandidate(currentSyntax, nextSyntax) &&
            TryAbsorbDefinitionLazyParagraphIntoLastListItem(currentList, nextParagraph, nextSyntax, sourceLines, options, state)) {
            combinedSyntax = CombineDefinitionLazyListParagraphSyntax(currentSyntax, nextSyntax, currentList);
            return true;
        }

        if (currentBlock is IMarkdownListBlock currentListWithTable &&
            nextBlock is TableBlock nextTable &&
            nextSyntax.Kind == MarkdownSyntaxKind.Table &&
            IsDefinitionLazyNestedBlockContinuationCandidate(currentSyntax, nextSyntax) &&
            TryAbsorbDefinitionLazyBlockIntoLastListItem(currentListWithTable, nextTable)) {
            combinedSyntax = CombineDefinitionLazyListNestedBlockSyntax(currentSyntax, nextSyntax, currentListWithTable);
            return true;
        }

        if (IsDefinitionLazyListContinuationCandidate(currentSyntax, nextSyntax) &&
            currentBlock is UnorderedListBlock currentUnordered &&
            nextBlock is UnorderedListBlock nextUnordered &&
            string.Equals(GetFirstListMarkerLiteral(currentSyntax), GetFirstListMarkerLiteral(nextSyntax), StringComparison.Ordinal)) {
            currentUnordered.Items.AddRange(nextUnordered.Items);
            combinedSyntax = CombineDefinitionLazyListSyntax(currentSyntax, nextSyntax, currentUnordered);
            return true;
        }

        if (IsDefinitionLazyListContinuationCandidate(currentSyntax, nextSyntax) &&
            currentBlock is OrderedListBlock currentOrdered &&
            nextBlock is OrderedListBlock nextOrdered &&
            string.Equals(
                GetFirstOrderedListMarkerDelimiter(currentSyntax),
                GetFirstOrderedListMarkerDelimiter(nextSyntax),
                StringComparison.Ordinal)) {
            currentOrdered.Items.AddRange(nextOrdered.Items);
            combinedSyntax = CombineDefinitionLazyListSyntax(currentSyntax, nextSyntax, currentOrdered);
            return true;
        }

        return false;
    }

    private static bool IsDefinitionLazyListContinuationCandidate(MarkdownSyntaxNode currentSyntax, MarkdownSyntaxNode nextSyntax) {
        if (currentSyntax == null ||
            nextSyntax == null ||
            currentSyntax.Kind != nextSyntax.Kind ||
            !currentSyntax.SourceSpan.HasValue ||
            !nextSyntax.SourceSpan.HasValue) {
            return false;
        }

        var currentSpan = currentSyntax.SourceSpan.Value;
        var nextSpan = nextSyntax.SourceSpan.Value;
        return currentSpan.EndLine + 1 == nextSpan.StartLine &&
            currentSpan.StartColumn.HasValue &&
            nextSpan.StartColumn.HasValue &&
            currentSpan.StartColumn.Value > nextSpan.StartColumn.Value;
    }

    private static bool IsDefinitionLazyParagraphContinuationCandidate(MarkdownSyntaxNode currentSyntax, MarkdownSyntaxNode nextSyntax) {
        if (currentSyntax == null ||
            nextSyntax == null ||
            (currentSyntax.Kind != MarkdownSyntaxKind.UnorderedList &&
             currentSyntax.Kind != MarkdownSyntaxKind.OrderedList) ||
            nextSyntax.Kind != MarkdownSyntaxKind.Paragraph ||
            !currentSyntax.SourceSpan.HasValue ||
            !nextSyntax.SourceSpan.HasValue) {
            return false;
        }

        var currentSpan = currentSyntax.SourceSpan.Value;
        var nextSpan = nextSyntax.SourceSpan.Value;
        return currentSpan.EndLine + 1 == nextSpan.StartLine &&
            currentSpan.StartColumn.HasValue &&
            nextSpan.StartColumn.HasValue &&
            currentSpan.StartColumn.Value > nextSpan.StartColumn.Value;
    }

    private static bool IsDefinitionLazyNestedBlockContinuationCandidate(MarkdownSyntaxNode currentSyntax, MarkdownSyntaxNode nextSyntax) {
        if (currentSyntax == null ||
            nextSyntax == null ||
            (currentSyntax.Kind != MarkdownSyntaxKind.UnorderedList &&
             currentSyntax.Kind != MarkdownSyntaxKind.OrderedList) ||
            !currentSyntax.SourceSpan.HasValue ||
            !nextSyntax.SourceSpan.HasValue) {
            return false;
        }

        var currentSpan = currentSyntax.SourceSpan.Value;
        var nextSpan = nextSyntax.SourceSpan.Value;
        return currentSpan.EndLine + 1 == nextSpan.StartLine &&
            currentSpan.StartColumn.HasValue &&
            nextSpan.StartColumn.HasValue &&
            currentSpan.StartColumn.Value > nextSpan.StartColumn.Value;
    }

    private static bool TryAbsorbDefinitionLazyParagraphIntoLastListItem(
        IMarkdownListBlock listBlock,
        ParagraphBlock paragraph,
        MarkdownSyntaxNode paragraphSyntax,
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (listBlock == null || paragraph == null || listBlock.ListItems.Count == 0) {
            return false;
        }

        var item = listBlock.ListItems[listBlock.ListItems.Count - 1];
        var continuationNodes = GetDefinitionLazyParagraphContinuationNodes(paragraph, paragraphSyntax, sourceLines, options, state);
        var nodes = new List<IMarkdownInline>(item.Content.Nodes.Count + continuationNodes.Count + 1);
        nodes.AddRange(item.Content.Nodes);
        if (item.Content.Nodes.Count > 0 && continuationNodes.Count > 0) {
            nodes.Add(new TextRun("\n"));
        }

        nodes.AddRange(continuationNodes);
        item.Content.AutoSpacing = false;
        item.Content.ReplaceItems(nodes);
        item.DefinitionLazyParagraphTailContinuation = continuationNodes.Count == 1 &&
            continuationNodes[0] is TextRun textRun &&
            textRun.Text.IndexOf('\n') >= 0;
        return true;
    }

    private static bool TryAbsorbDefinitionLazyBlockIntoLastListItem(
        IMarkdownListBlock listBlock,
        IMarkdownBlock block) {
        if (listBlock == null || block == null || listBlock.ListItems.Count == 0) {
            return false;
        }

        var item = listBlock.ListItems[listBlock.ListItems.Count - 1];
        item.NestedBlocks.Add(block);
        return true;
    }

    private static IReadOnlyList<IMarkdownInline> GetDefinitionLazyParagraphContinuationNodes(
        ParagraphBlock paragraph,
        MarkdownSyntaxNode paragraphSyntax,
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (TryGetDefinitionLazyParagraphSourceText(paragraphSyntax, sourceLines, out var sourceText)) {
            if (!ContainsBackslashEscapableCharacter(sourceText)) {
                return new IMarkdownInline[] {
                    new TextRun(sourceText)
                };
            }

            return ParseDefinitionLazyListItemContinuationNodes(sourceText, options, state);
        }

        return paragraph.Inlines.Nodes;
    }

    private static IReadOnlyList<IMarkdownInline> ParseDefinitionLazyListItemContinuationNodes(
        string sourceText,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (string.IsNullOrEmpty(sourceText)) {
            return Array.Empty<IMarkdownInline>();
        }

        var nodes = new List<IMarkdownInline>();
        var lines = sourceText.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++) {
            if (lineIndex > 0) {
                nodes.Add(new SoftBreakInline());
            }

            var parsed = ParseInlines(lines[lineIndex], options, state);
            for (int nodeIndex = 0; nodeIndex < parsed.Nodes.Count; nodeIndex++) {
                nodes.Add(parsed.Nodes[nodeIndex]);
            }
        }

        return nodes;
    }

    private static bool TryGetDefinitionLazyParagraphSourceText(
        MarkdownSyntaxNode paragraphSyntax,
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        out string sourceText) {
        sourceText = string.Empty;
        if (paragraphSyntax == null ||
            !paragraphSyntax.SourceSpan.HasValue ||
            sourceLines == null ||
            sourceLines.Count == 0) {
            return false;
        }

        var span = paragraphSyntax.SourceSpan.Value;
        var lines = new List<string>();
        for (int i = 0; i < sourceLines.Count; i++) {
            var line = sourceLines[i];
            if (line.AbsoluteLine < span.StartLine || line.AbsoluteLine > span.EndLine) {
                continue;
            }

            lines.Add(line.Text);
        }

        if (lines.Count <= 1) {
            return false;
        }

        sourceText = string.Join("\n", lines);
        return sourceText.Length > 0;
    }

    private static MarkdownSyntaxNode CombineDefinitionLazyListParagraphSyntax(
        MarkdownSyntaxNode currentSyntax,
        MarkdownSyntaxNode nextSyntax,
        IMarkdownBlock associatedBlock) {
        var children = new List<MarkdownSyntaxNode>(currentSyntax.Children);
        for (int i = children.Count - 1; i >= 0; i--) {
            if (children[i].Kind != MarkdownSyntaxKind.ListItem) {
                continue;
            }

            children[i] = CombineDefinitionLazyListItemParagraphSyntax(children[i], nextSyntax);
            break;
        }

        return new MarkdownSyntaxNode(
            currentSyntax.Kind,
            CombineDefinitionLazyListSourceSpans(currentSyntax.SourceSpan, nextSyntax.SourceSpan),
            currentSyntax.Literal,
            children,
            associatedBlock,
            currentSyntax.CustomKind,
            currentSyntax.Attributes);
    }

    private static MarkdownSyntaxNode CombineDefinitionLazyListItemParagraphSyntax(MarkdownSyntaxNode listItemSyntax, MarkdownSyntaxNode nextSyntax) {
        var children = new List<MarkdownSyntaxNode>(listItemSyntax.Children);
        for (int i = children.Count - 1; i >= 0; i--) {
            if (children[i].Kind != MarkdownSyntaxKind.Paragraph) {
                continue;
            }

            children[i] = CombineDefinitionLazyParagraphSyntax(children[i], nextSyntax);
            break;
        }

        return new MarkdownSyntaxNode(
            listItemSyntax.Kind,
            CombineDefinitionLazyListSourceSpans(listItemSyntax.SourceSpan, nextSyntax.SourceSpan),
            listItemSyntax.Literal,
            children,
            listItemSyntax.AssociatedObject,
            listItemSyntax.CustomKind,
            listItemSyntax.Attributes);
    }

    private static MarkdownSyntaxNode CombineDefinitionLazyParagraphSyntax(MarkdownSyntaxNode paragraphSyntax, MarkdownSyntaxNode nextSyntax) {
        var children = new List<MarkdownSyntaxNode>(paragraphSyntax.Children.Count + nextSyntax.Children.Count);
        children.AddRange(paragraphSyntax.Children);
        children.AddRange(nextSyntax.Children);
        var literal = string.IsNullOrEmpty(paragraphSyntax.Literal)
            ? nextSyntax.Literal
            : string.IsNullOrEmpty(nextSyntax.Literal)
                ? paragraphSyntax.Literal
                : paragraphSyntax.Literal + "\n" + nextSyntax.Literal;

        return new MarkdownSyntaxNode(
            paragraphSyntax.Kind,
            CombineDefinitionLazyListSourceSpans(paragraphSyntax.SourceSpan, nextSyntax.SourceSpan),
            literal,
            children,
            paragraphSyntax.AssociatedObject,
            paragraphSyntax.CustomKind,
            paragraphSyntax.Attributes);
    }

    private static MarkdownSyntaxNode CombineDefinitionLazyListNestedBlockSyntax(
        MarkdownSyntaxNode currentSyntax,
        MarkdownSyntaxNode nextSyntax,
        IMarkdownBlock associatedBlock) {
        var children = new List<MarkdownSyntaxNode>(currentSyntax.Children);
        for (int i = children.Count - 1; i >= 0; i--) {
            if (children[i].Kind != MarkdownSyntaxKind.ListItem) {
                continue;
            }

            children[i] = CombineDefinitionLazyListItemNestedBlockSyntax(children[i], nextSyntax);
            break;
        }

        return new MarkdownSyntaxNode(
            currentSyntax.Kind,
            CombineDefinitionLazyListSourceSpans(currentSyntax.SourceSpan, nextSyntax.SourceSpan),
            currentSyntax.Literal,
            children,
            associatedBlock,
            currentSyntax.CustomKind,
            currentSyntax.Attributes);
    }

    private static MarkdownSyntaxNode CombineDefinitionLazyListItemNestedBlockSyntax(MarkdownSyntaxNode listItemSyntax, MarkdownSyntaxNode nextSyntax) {
        var children = new List<MarkdownSyntaxNode>(listItemSyntax.Children.Count + 1);
        children.AddRange(listItemSyntax.Children);
        children.Add(nextSyntax);

        return new MarkdownSyntaxNode(
            listItemSyntax.Kind,
            CombineDefinitionLazyListSourceSpans(listItemSyntax.SourceSpan, nextSyntax.SourceSpan),
            listItemSyntax.Literal,
            children,
            listItemSyntax.AssociatedObject,
            listItemSyntax.CustomKind,
            listItemSyntax.Attributes);
    }

    private static MarkdownSyntaxNode CombineDefinitionLazyListSyntax(
        MarkdownSyntaxNode currentSyntax,
        MarkdownSyntaxNode nextSyntax,
        IMarkdownBlock associatedBlock) {
        var children = new List<MarkdownSyntaxNode>(currentSyntax.Children.Count + nextSyntax.Children.Count);
        children.AddRange(currentSyntax.Children);
        children.AddRange(nextSyntax.Children);
        return new MarkdownSyntaxNode(
            currentSyntax.Kind,
            CombineDefinitionLazyListSourceSpans(currentSyntax.SourceSpan, nextSyntax.SourceSpan),
            currentSyntax.Literal,
            children,
            associatedBlock,
            currentSyntax.CustomKind,
            currentSyntax.Attributes);
    }

    private static MarkdownSourceSpan? CombineDefinitionLazyListSourceSpans(MarkdownSourceSpan? first, MarkdownSourceSpan? second) {
        if (!first.HasValue) {
            return second;
        }

        if (!second.HasValue) {
            return first;
        }

        var start = first.Value;
        var end = second.Value;
        if (!start.StartColumn.HasValue || !end.EndColumn.HasValue) {
            return new MarkdownSourceSpan(start.StartLine, end.EndLine);
        }

        return new MarkdownSourceSpan(
            start.StartLine,
            start.StartColumn.Value,
            end.EndLine,
            end.EndColumn.Value);
    }

    private static string? GetFirstListMarkerLiteral(MarkdownSyntaxNode listSyntax) {
        if (listSyntax == null) {
            return null;
        }

        foreach (var listItem in listSyntax.Children) {
            if (listItem.Kind != MarkdownSyntaxKind.ListItem) {
                continue;
            }

            foreach (var child in listItem.Children) {
                if (child.Kind == MarkdownSyntaxKind.ListMarker) {
                    return child.Literal;
                }
            }
        }

        return null;
    }

    private static string? GetFirstOrderedListMarkerDelimiter(MarkdownSyntaxNode listSyntax) {
        var literal = GetFirstListMarkerLiteral(listSyntax);
        if (string.IsNullOrEmpty(literal)) {
            return null;
        }

        char delimiter = literal![literal.Length - 1];
        return delimiter == '.' || delimiter == ')' ? delimiter.ToString() : null;
    }
}
