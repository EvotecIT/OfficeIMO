using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    internal static IReadOnlyList<OfficeDocumentContentItem> Content(OfficeDocumentReadResult document) {
        var items = new List<OfficeDocumentContentItem>();
        var hints = new Dictionary<OfficeDocumentContentItem, long>();
        var anchors = new Dictionary<string, (ReaderLocation Location, long Order)?>(StringComparer.Ordinal);
        void Add(OfficeDocumentContentItem item, long? order = null) {
            hints.Add(item, order ?? (long)items.Count * 2);
            items.Add(item);
        }
        foreach (OfficeDocumentBlock block in Blocks(document)) {
            long order = (long)items.Count * 2;
            Add(new(block, null, null, block.Location));
            string? anchor = block.Location?.BlockAnchor;
            if (!string.IsNullOrWhiteSpace(anchor)) {
                // An ambiguous anchor cannot supply a reliable position for a table.
                if (anchors.ContainsKey(anchor!)) anchors[anchor!] = null;
                else anchors.Add(anchor!, (block.Location!, order));
            }
        }
        if (!items.Any(item => !string.IsNullOrWhiteSpace(item.Block?.Text))) {
            foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>())
                if (chunk != null) Add(new(null, null, chunk, chunk.Location));
        }
        var chunkLocations = new Dictionary<ReaderTable, ReaderLocation>(ReferenceIdentityComparer<ReaderTable>.Instance);
        foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>()) {
            if (chunk?.Location == null) continue;
            foreach (ReaderTable table in chunk.Tables ?? Array.Empty<ReaderTable>())
                if (table != null && !chunkLocations.ContainsKey(table)) chunkLocations.Add(table, chunk.Location);
        }
        foreach (ReaderTable table in Tables(document)) {
            ReaderLocation? location = table.Location;
            long? order = null;
            if (chunkLocations.TryGetValue(table, out ReaderLocation? chunkLocation))
                location = MergeLocation(location, chunkLocation, location?.TableIndex);
            if (!string.IsNullOrWhiteSpace(location?.BlockAnchor) && anchors.TryGetValue(location!.BlockAnchor!, out var anchor)
                && anchor.HasValue && SameContainerWhenKnown(location, anchor.Value.Location)) {
                location = MergeLocation(location, anchor.Value.Location, location?.TableIndex);
                order = anchor.Value.Order + 1;
            }
            Add(new(null, table, null, location), order);
        }
        return OrderSourceItems(items, item => item.Location, document.Pages, item => hints[item]);
    }

    private static bool SameContainerWhenKnown(ReaderLocation left, ReaderLocation right) =>
        (string.IsNullOrWhiteSpace(left.Path) || string.IsNullOrWhiteSpace(right.Path) || left.Path == right.Path)
        && (!left.Page.HasValue || !right.Page.HasValue || left.Page == right.Page)
        && (!left.Slide.HasValue || !right.Slide.HasValue || left.Slide == right.Slide)
        && (string.IsNullOrWhiteSpace(left.Sheet) || string.IsNullOrWhiteSpace(right.Sheet) || left.Sheet == right.Sheet);

    private static IReadOnlyList<T> OrderSourceItems<T>(IEnumerable<T> candidates, Func<T, ReaderLocation?> locationSelector,
        IEnumerable<OfficeDocumentPage>? pages = null, Func<T, long>? encounterOrder = null) {
        var sheetOrder = new Dictionary<string, int>(StringComparer.Ordinal);
        void RegisterSheet(string? sheet) {
            if (!string.IsNullOrWhiteSpace(sheet) && !sheetOrder.ContainsKey(sheet!)) sheetOrder.Add(sheet!, sheetOrder.Count);
        }
        foreach (OfficeDocumentPage page in pages ?? Array.Empty<OfficeDocumentPage>()) RegisterSheet(page?.Location?.Sheet);
        var ordered = candidates.Select((item, index) => (Item: item, Location: locationSelector(item), Index: index)).ToList();
        foreach (var item in ordered) RegisterSheet(item.Location?.Sheet);
        ordered.Sort((left, right) => {
            int comparison = string.CompareOrdinal(BuildContainerOrderKey(left.Location, sheetOrder), BuildContainerOrderKey(right.Location, sheetOrder));
            if (comparison != 0) return comparison;
            comparison = BuildBlockPosition(left.Location).CompareTo(BuildBlockPosition(right.Location));
            if (comparison != 0) return comparison;
            comparison = (encounterOrder?.Invoke(left.Item) ?? left.Index).CompareTo(encounterOrder?.Invoke(right.Item) ?? right.Index);
            return comparison != 0 ? comparison : left.Index.CompareTo(right.Index);
        });
        return ordered.Select(item => item.Item).ToArray();
    }
}
