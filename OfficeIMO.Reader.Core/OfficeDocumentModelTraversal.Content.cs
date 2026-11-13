namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    internal static IReadOnlyList<OfficeDocumentContentItem> Content(OfficeDocumentReadResult document) {
        var items = new List<OfficeDocumentContentItem>();
        var hints = new Dictionary<OfficeDocumentContentItem, long>();
        var anchors = new Dictionary<string, List<(ReaderLocation Location, long Order)>>(StringComparer.Ordinal);
        void Add(OfficeDocumentContentItem item, long? order = null) {
            hints.Add(item, order ?? (long)items.Count * 2);
            items.Add(item);
        }
        foreach (OfficeDocumentBlock block in Blocks(document)) {
            long order = (long)items.Count * 2;
            Add(new(block, null, null, block.Location));
            string? anchor = block.Location?.BlockAnchor;
            if (!string.IsNullOrWhiteSpace(anchor)) {
                if (!anchors.TryGetValue(anchor!, out var positions)) anchors.Add(anchor!, positions = new());
                positions.Add((block.Location!, order));
            }
        }
        if (!items.Any(item => !string.IsNullOrWhiteSpace(item.Block?.Text))) {
            foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>())
                if (chunk != null) Add(new(null, null, chunk, chunk.Location));
        }
        foreach (ReaderTable table in Tables(document)) {
            ReaderLocation? location = table.Location;
            long? order = null;
            if (!string.IsNullOrWhiteSpace(location?.BlockAnchor) && anchors.TryGetValue(location!.BlockAnchor!, out var positions)) {
                // Readers can reuse a local anchor on different pages, slides, sheets or paths.
                // Only ambiguity among compatible containers prevents a reliable position.
                var matches = positions.Where(position => SameContainerWhenKnown(location, position.Location)).Take(2).ToArray();
                if (matches.Length == 1) {
                    location = MergeLocation(location, matches[0].Location, location?.TableIndex);
                    order = matches[0].Order + 1;
                }
            }
            Add(new(null, table, null, location), order);
        }
        return OrderSourceItems(items, item => item.Location, document.Pages, item => hints[item]);
    }

    /// <summary>Matches source containers without treating absent provenance as a conflicting container.</summary>
    internal static bool SameContainerWhenKnown(ReaderLocation left, ReaderLocation right) =>
        (string.IsNullOrWhiteSpace(left.Path) || string.IsNullOrWhiteSpace(right.Path) || left.Path == right.Path)
        && (!left.Page.HasValue || !right.Page.HasValue || left.Page == right.Page)
        && (!left.Slide.HasValue || !right.Slide.HasValue || left.Slide == right.Slide)
        && (string.IsNullOrWhiteSpace(left.Sheet) || string.IsNullOrWhiteSpace(right.Sheet) || left.Sheet == right.Sheet);

    private readonly struct SourceOrderKey<T> {
        internal SourceOrderKey(T item, int pathIndex, long logicalOrder, string containerKey, int blockPosition, long encounter, int index) {
            Item = item; PathIndex = pathIndex; LogicalOrder = logicalOrder; ContainerKey = containerKey;
            BlockPosition = blockPosition; Encounter = encounter; Index = index;
        }
        internal T Item { get; }
        internal int PathIndex { get; }
        internal long LogicalOrder { get; }
        internal string ContainerKey { get; }
        internal int BlockPosition { get; }
        internal long Encounter { get; }
        internal int Index { get; }
    }

    private static IReadOnlyList<T> OrderSourceItems<T>(IEnumerable<T> candidates, Func<T, ReaderLocation?> locationSelector,
        IEnumerable<OfficeDocumentPage>? pages = null, Func<T, long>? encounterOrder = null) {
        var pathOrder = new Dictionary<string, int>(StringComparer.Ordinal);
        void RegisterPath(string? path) {
            if (!string.IsNullOrWhiteSpace(path) && !pathOrder.ContainsKey(path!)) pathOrder.Add(path!, pathOrder.Count);
        }
        var sheetOrder = new Dictionary<string, int>(StringComparer.Ordinal);
        void RegisterSheet(string? sheet) {
            if (!string.IsNullOrWhiteSpace(sheet) && !sheetOrder.ContainsKey(sheet!)) sheetOrder.Add(sheet!, sheetOrder.Count);
        }
        foreach (OfficeDocumentPage page in pages ?? Array.Empty<OfficeDocumentPage>()) {
            ReaderLocation? location = page?.GetResolvedLocation();
            RegisterPath(location?.Path);
            RegisterSheet(location?.Sheet);
        }
        var ordered = candidates.Select((item, index) => (Item: item, Location: locationSelector(item), Index: index)).ToList();
        foreach (var item in ordered) {
            RegisterPath(item.Location?.Path);
            RegisterSheet(item.Location?.Sheet);
        }
        int PathIndex(ReaderLocation? location) => !string.IsNullOrWhiteSpace(location?.Path) ? pathOrder[location!.Path!] : int.MaxValue;
        // Each item's keys are computed once; the comparison then reads them instead of rebuilding container keys
        // and looking up encounter hints on every comparison. The keys and their order are unchanged.
        var keyed = new List<SourceOrderKey<T>>(ordered.Count);
        foreach (var item in ordered) {
            keyed.Add(new SourceOrderKey<T>(item.Item, PathIndex(item.Location), item.Location?.LogicalOrder ?? long.MaxValue,
                BuildContainerOrderKey(item.Location, sheetOrder), BuildBlockPosition(item.Location),
                encounterOrder?.Invoke(item.Item) ?? item.Index, item.Index));
        }
        keyed.Sort(static (left, right) => {
            int comparison = left.PathIndex.CompareTo(right.PathIndex);
            if (comparison != 0) return comparison;
            comparison = left.LogicalOrder.CompareTo(right.LogicalOrder);
            if (comparison != 0) return comparison;
            comparison = string.CompareOrdinal(left.ContainerKey, right.ContainerKey);
            if (comparison != 0) return comparison;
            comparison = left.BlockPosition.CompareTo(right.BlockPosition);
            if (comparison != 0) return comparison;
            comparison = left.Encounter.CompareTo(right.Encounter);
            return comparison != 0 ? comparison : left.Index.CompareTo(right.Index);
        });
        var result = new T[keyed.Count];
        for (int index = 0; index < keyed.Count; index++) result[index] = keyed[index].Item;
        return result;
    }
}
