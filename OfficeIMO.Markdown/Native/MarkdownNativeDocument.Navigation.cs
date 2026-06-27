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

    /// <summary>Enumerates all native definition-list groups in document order, including nested definition lists.</summary>
    public IEnumerable<MarkdownNativeDefinitionListGroup> EnumerateDefinitionListGroups() {
        for (var i = 0; i < Blocks.Count; i++) {
            foreach (var group in EnumerateDefinitionListGroups(Blocks[i])) {
                yield return group;
            }
        }
    }

    /// <summary>Finds the deepest native definition-list group whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeDefinitionListGroup? FindDefinitionListGroupAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindDefinitionListGroupAtPosition(Blocks[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Enumerates all native definition-list terms in document order, including nested definition lists.</summary>
    public IEnumerable<MarkdownNativeDefinitionListTerm> EnumerateDefinitionListTerms() {
        for (var i = 0; i < Blocks.Count; i++) {
            foreach (var term in EnumerateDefinitionListTerms(Blocks[i])) {
                yield return term;
            }
        }
    }

    /// <summary>Finds the native definition-list term whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeDefinitionListTerm? FindDefinitionListTermAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindDefinitionListTermAtPosition(Blocks[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Enumerates all native definition-list definitions in document order, including nested definition lists.</summary>
    public IEnumerable<MarkdownNativeDefinitionListDefinition> EnumerateDefinitionListDefinitions() {
        for (var i = 0; i < Blocks.Count; i++) {
            foreach (var definition in EnumerateDefinitionListDefinitions(Blocks[i])) {
                yield return definition;
            }
        }
    }

    /// <summary>Finds the deepest native definition-list definition whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeDefinitionListDefinition? FindDefinitionListDefinitionAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindDefinitionListDefinitionAtPosition(Blocks[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Enumerates all source-backed native inline metadata leaves in document order.</summary>
    public IEnumerable<MarkdownNativeInlineMetadata> EnumerateInlineMetadata() {
        foreach (var inline in EnumerateInlines()) {
            for (var i = 0; i < inline.Metadata.Count; i++) {
                yield return inline.Metadata[i];
            }
        }
    }

    /// <summary>Enumerates source-backed native inline metadata leaves with the supplied metadata name in document order.</summary>
    public IEnumerable<MarkdownNativeInlineMetadata> EnumerateInlineMetadata(string name) {
        if (string.IsNullOrWhiteSpace(name)) {
            yield break;
        }

        foreach (var metadata in EnumerateInlineMetadata()) {
            if (string.Equals(metadata.Name, name, StringComparison.OrdinalIgnoreCase)) {
                yield return metadata;
            }
        }
    }

    /// <summary>Finds the first native inline metadata leaf whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeInlineMetadata? FindInlineMetadataAtPosition(int lineNumber, int columnNumber) {
        foreach (var metadata in EnumerateInlineMetadata()) {
            if (ContainsPosition(metadata.SourceSpan, lineNumber, columnNumber)) {
                return metadata;
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

    private static IEnumerable<MarkdownNativeDefinitionListGroup> EnumerateDefinitionListGroups(MarkdownNativeBlock block) {
        if (block is MarkdownNativeDefinitionListBlock definitionList) {
            for (var i = 0; i < definitionList.Groups.Count; i++) {
                yield return definitionList.Groups[i];
            }
        }

        foreach (var child in GetChildBlocks(block)) {
            foreach (var group in EnumerateDefinitionListGroups(child)) {
                yield return group;
            }
        }
    }

    private static MarkdownNativeDefinitionListGroup? FindDefinitionListGroupAtPosition(MarkdownNativeBlock block, int lineNumber, int columnNumber) {
        if (block is MarkdownNativeDefinitionListBlock definitionList) {
            for (var i = 0; i < definitionList.Groups.Count; i++) {
                var group = definitionList.Groups[i];
                for (var definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                    var definition = group.Definitions[definitionIndex];
                    for (var childIndex = 0; childIndex < definition.Children.Count; childIndex++) {
                        var childMatch = FindDefinitionListGroupAtPosition(definition.Children[childIndex], lineNumber, columnNumber);
                        if (childMatch != null) {
                            return childMatch;
                        }
                    }
                }

                if (ContainsPosition(group.SourceSpan, lineNumber, columnNumber)) {
                    return group;
                }
            }

            return null;
        }

        foreach (var child in GetChildBlocks(block)) {
            var match = FindDefinitionListGroupAtPosition(child, lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static IEnumerable<MarkdownNativeDefinitionListTerm> EnumerateDefinitionListTerms(MarkdownNativeBlock block) {
        if (block is MarkdownNativeDefinitionListBlock definitionList) {
            for (var groupIndex = 0; groupIndex < definitionList.Groups.Count; groupIndex++) {
                var group = definitionList.Groups[groupIndex];
                for (var termIndex = 0; termIndex < group.Terms.Count; termIndex++) {
                    yield return group.Terms[termIndex];
                }
            }
        }

        foreach (var child in GetChildBlocks(block)) {
            foreach (var term in EnumerateDefinitionListTerms(child)) {
                yield return term;
            }
        }
    }

    private static MarkdownNativeDefinitionListTerm? FindDefinitionListTermAtPosition(MarkdownNativeBlock block, int lineNumber, int columnNumber) {
        if (block is MarkdownNativeDefinitionListBlock definitionList) {
            for (var groupIndex = 0; groupIndex < definitionList.Groups.Count; groupIndex++) {
                var group = definitionList.Groups[groupIndex];
                for (var termIndex = 0; termIndex < group.Terms.Count; termIndex++) {
                    if (ContainsPosition(group.Terms[termIndex].SourceSpan, lineNumber, columnNumber)) {
                        return group.Terms[termIndex];
                    }
                }

                for (var definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                    var definition = group.Definitions[definitionIndex];
                    for (var childIndex = 0; childIndex < definition.Children.Count; childIndex++) {
                        var childMatch = FindDefinitionListTermAtPosition(definition.Children[childIndex], lineNumber, columnNumber);
                        if (childMatch != null) {
                            return childMatch;
                        }
                    }
                }
            }

            return null;
        }

        foreach (var child in GetChildBlocks(block)) {
            var match = FindDefinitionListTermAtPosition(child, lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static IEnumerable<MarkdownNativeDefinitionListDefinition> EnumerateDefinitionListDefinitions(MarkdownNativeBlock block) {
        if (block is MarkdownNativeDefinitionListBlock definitionList) {
            for (var groupIndex = 0; groupIndex < definitionList.Groups.Count; groupIndex++) {
                var group = definitionList.Groups[groupIndex];
                for (var definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                    yield return group.Definitions[definitionIndex];
                }
            }
        }

        foreach (var child in GetChildBlocks(block)) {
            foreach (var definition in EnumerateDefinitionListDefinitions(child)) {
                yield return definition;
            }
        }
    }

    private static MarkdownNativeDefinitionListDefinition? FindDefinitionListDefinitionAtPosition(MarkdownNativeBlock block, int lineNumber, int columnNumber) {
        if (block is MarkdownNativeDefinitionListBlock definitionList) {
            for (var groupIndex = 0; groupIndex < definitionList.Groups.Count; groupIndex++) {
                var group = definitionList.Groups[groupIndex];
                for (var definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                    var definition = group.Definitions[definitionIndex];
                    for (var childIndex = 0; childIndex < definition.Children.Count; childIndex++) {
                        var childMatch = FindDefinitionListDefinitionAtPosition(definition.Children[childIndex], lineNumber, columnNumber);
                        if (childMatch != null) {
                            return childMatch;
                        }
                    }

                    if (ContainsPosition(definition.SourceSpan, lineNumber, columnNumber)) {
                        return definition;
                    }
                }
            }

            return null;
        }

        foreach (var child in GetChildBlocks(block)) {
            var match = FindDefinitionListDefinitionAtPosition(child, lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static bool ContainsPosition(MarkdownSourceSpan? sourceSpan, int lineNumber, int columnNumber) =>
        sourceSpan.HasValue && sourceSpan.Value.ContainsPosition(lineNumber, columnNumber);
}
