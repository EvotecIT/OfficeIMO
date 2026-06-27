namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    /// <summary>Enumerates all native list items in document order, including nested list items.</summary>
    public IEnumerable<MarkdownNativeListItem> EnumerateListItems() {
        for (var i = 0; i < Blocks.Count; i++) {
            foreach (var item in EnumerateListItems(Blocks[i])) {
                yield return item;
            }
        }
    }

    /// <summary>Finds the deepest native list item whose content or marker source span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeListItem? FindListItemAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindListItemAtPosition(Blocks[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Finds a native list item by deterministic id.</summary>
    public MarkdownNativeListItem? FindListItemById(string id) {
        if (string.IsNullOrWhiteSpace(id)) {
            return null;
        }

        foreach (var item in EnumerateListItems()) {
            if (string.Equals(item.Id, id, StringComparison.Ordinal)) {
                return item;
            }
        }

        return null;
    }

    /// <summary>Enumerates all native table cells in document order, including cells in nested tables.</summary>
    public IEnumerable<MarkdownNativeTableCell> EnumerateTableCells() {
        for (var i = 0; i < Blocks.Count; i++) {
            foreach (var cell in EnumerateTableCells(Blocks[i])) {
                yield return cell;
            }
        }
    }

    /// <summary>Finds the deepest native table cell whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeTableCell? FindTableCellAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindTableCellAtPosition(Blocks[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static IEnumerable<MarkdownNativeListItem> EnumerateListItems(MarkdownNativeBlock block) {
        if (block is MarkdownNativeListBlock list) {
            for (var i = 0; i < list.Items.Count; i++) {
                yield return list.Items[i];
                for (var j = 0; j < list.Items[i].Children.Count; j++) {
                    foreach (var child in EnumerateListItems(list.Items[i].Children[j])) {
                        yield return child;
                    }
                }
            }

            yield break;
        }

        foreach (var child in GetChildBlocks(block)) {
            foreach (var item in EnumerateListItems(child)) {
                yield return item;
            }
        }
    }

    private static MarkdownNativeListItem? FindListItemAtPosition(MarkdownNativeBlock block, int lineNumber, int columnNumber) {
        if (block is MarkdownNativeListBlock list) {
            for (var i = 0; i < list.Items.Count; i++) {
                var item = list.Items[i];
                for (var j = 0; j < item.Children.Count; j++) {
                    var childMatch = FindListItemAtPosition(item.Children[j], lineNumber, columnNumber);
                    if (childMatch != null) {
                        return childMatch;
                    }
                }

                if (ContainsPosition(item, lineNumber, columnNumber)) {
                    return item;
                }
            }

            return null;
        }

        foreach (var child in GetChildBlocks(block)) {
            var match = FindListItemAtPosition(child, lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static bool ContainsPosition(MarkdownNativeListItem item, int lineNumber, int columnNumber) =>
        ContainsPosition(item.SourceSpan, lineNumber, columnNumber)
        || ContainsPosition(item.MarkerSourceSpan, lineNumber, columnNumber)
        || ContainsPosition(item.TaskMarkerSourceSpan, lineNumber, columnNumber);

    private static IEnumerable<MarkdownNativeTableCell> EnumerateTableCells(MarkdownNativeBlock block) {
        if (block is MarkdownNativeTableBlock table) {
            for (var i = 0; i < table.HeaderCells.Count; i++) {
                yield return table.HeaderCells[i];
                foreach (var child in EnumerateTableCells(table.HeaderCells[i])) {
                    yield return child;
                }
            }

            for (var row = 0; row < table.Rows.Count; row++) {
                for (var column = 0; column < table.Rows[row].Count; column++) {
                    yield return table.Rows[row][column];
                    foreach (var child in EnumerateTableCells(table.Rows[row][column])) {
                        yield return child;
                    }
                }
            }

            yield break;
        }

        foreach (var child in GetChildBlocks(block)) {
            foreach (var cell in EnumerateTableCells(child)) {
                yield return cell;
            }
        }
    }

    private static IEnumerable<MarkdownNativeTableCell> EnumerateTableCells(MarkdownNativeTableCell cell) {
        for (var i = 0; i < cell.Children.Count; i++) {
            foreach (var child in EnumerateTableCells(cell.Children[i])) {
                yield return child;
            }
        }
    }

    private static MarkdownNativeTableCell? FindTableCellAtPosition(MarkdownNativeBlock block, int lineNumber, int columnNumber) {
        if (block is MarkdownNativeTableBlock table) {
            for (var i = 0; i < table.HeaderCells.Count; i++) {
                var childMatch = FindTableCellAtPosition(table.HeaderCells[i], lineNumber, columnNumber);
                if (childMatch != null) {
                    return childMatch;
                }

                if (ContainsPosition(table.HeaderCells[i].SourceSpan, lineNumber, columnNumber)) {
                    return table.HeaderCells[i];
                }
            }

            for (var row = 0; row < table.Rows.Count; row++) {
                for (var column = 0; column < table.Rows[row].Count; column++) {
                    var cell = table.Rows[row][column];
                    var childMatch = FindTableCellAtPosition(cell, lineNumber, columnNumber);
                    if (childMatch != null) {
                        return childMatch;
                    }

                    if (ContainsPosition(cell.SourceSpan, lineNumber, columnNumber)) {
                        return cell;
                    }
                }
            }

            return null;
        }

        foreach (var child in GetChildBlocks(block)) {
            var match = FindTableCellAtPosition(child, lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static MarkdownNativeTableCell? FindTableCellAtPosition(MarkdownNativeTableCell cell, int lineNumber, int columnNumber) {
        for (var i = 0; i < cell.Children.Count; i++) {
            var match = FindTableCellAtPosition(cell.Children[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static bool ContainsPosition(MarkdownSourceSpan? sourceSpan, int lineNumber, int columnNumber) =>
        sourceSpan.HasValue && sourceSpan.Value.ContainsPosition(lineNumber, columnNumber);
}
