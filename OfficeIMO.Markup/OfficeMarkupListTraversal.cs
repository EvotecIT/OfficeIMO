namespace OfficeIMO.Markup;

internal readonly struct OfficeMarkupListEntry {
    internal OfficeMarkupListEntry(
        OfficeMarkupListBlock sourceList,
        OfficeMarkupListItem item,
        int depth,
        int ordinal) {
        SourceList = sourceList;
        Item = item;
        Depth = depth;
        Ordinal = ordinal;
    }

    internal OfficeMarkupListBlock SourceList { get; }
    internal OfficeMarkupListItem Item { get; }
    internal int Depth { get; }
    internal int Ordinal { get; }
    internal string Marker => SourceList.Ordered
        ? Ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture) + "."
        : "-";
}

internal static class OfficeMarkupListTraversal {
    internal static IEnumerable<OfficeMarkupListEntry> Enumerate(OfficeMarkupListBlock list, int depth = 0) {
        for (int index = 0; index < list.Items.Count; index++) {
            var item = list.Items[index];
            yield return new OfficeMarkupListEntry(list, item, depth, list.Start + index);

            foreach (var nestedList in item.Blocks.OfType<OfficeMarkupListBlock>()) {
                foreach (var nestedItem in Enumerate(nestedList, depth + 1)) {
                    yield return nestedItem;
                }
            }
        }
    }
}
