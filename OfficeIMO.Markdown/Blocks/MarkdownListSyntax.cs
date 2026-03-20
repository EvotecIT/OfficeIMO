namespace OfficeIMO.Markdown;

internal static class MarkdownListSyntax {
    internal static MarkdownSyntaxNode BuildListBlockNode(IMarkdownListBlock listBlock, MarkdownSourceSpan? span) {
        var children = BuildListItemSyntaxNodes(listBlock.ListItems, listBlock.ListSyntaxKind);
        return new MarkdownSyntaxNode(
            listBlock.ListSyntaxKind,
            span ?? MarkdownBlockSyntaxBuilder.GetAggregateSpan(children),
            listBlock.ListLiteral,
            children,
            listBlock);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildListItemSyntaxNodes(IReadOnlyList<ListItem> items, MarkdownSyntaxKind listKind) {
        int index = 0;
        return BuildListItemSyntaxNodes(items, listKind, ref index, 0);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildListItemSyntaxNodes(IReadOnlyList<ListItem> items, MarkdownSyntaxKind listKind, ref int index, int level) {
        var nodes = new List<MarkdownSyntaxNode>();
        while (index < items.Count) {
            var item = items[index];
            if (item.Level < level) break;
            if (item.Level > level) {
                index++;
                continue;
            }

            index++;

            MarkdownSyntaxNode? nestedList = null;
            if (index < items.Count && items[index].Level > level) {
                var nestedLevel = items[index].Level;
                var nestedItems = BuildListItemSyntaxNodes(items, listKind, ref index, nestedLevel);
                nestedList = new MarkdownSyntaxNode(listKind, MarkdownBlockSyntaxBuilder.GetAggregateSpan(nestedItems), children: nestedItems);
            }

            nodes.Add(item.BuildSyntaxNode(nestedList));
        }

        return nodes;
    }
}
