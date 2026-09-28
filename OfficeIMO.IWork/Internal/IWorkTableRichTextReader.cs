namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves rich-text cell references through the table catalog and text-storage graph.</summary>
internal static class IWorkTableRichTextReader {
    private const uint DataListArchive = 6005;
    private const uint RichTextWrapperArchive = 6218;
    private const uint TextStorageArchive = 2001;

    internal static IReadOnlyDictionary<uint, string> Read(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkProjectionBudget projectionBudget,
        IWorkReadOptions options, int maximumEntries, out bool fullyReconstructed) {
        var strings = new Dictionary<uint, string>();
        var seenIdentifiers = new HashSet<uint>();
        fullyReconstructed = true;
        if (!store.HasField(17)) return strings;
        IWorkArchiveRecord? list = index.Dereference(store, 17);
        if (store.FieldCount(17) != 1 || list?.MessageType != DataListArchive) {
            fullyReconstructed = false;
            return strings;
        }
        if (!IWorkNumbersReader.TryGetCatalogEntryCount(list, maximumEntries, options,
                "rich-text", out int entryCount)) {
            fullyReconstructed = false;
            return strings;
        }
        projectionBudget.AddTableCatalogEntries(entryCount);
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            index.Message(list), 3, out bool malformedEntries);
        if (malformedEntries) fullyReconstructed = false;
        foreach (IWorkWireMessage entry in entries) {
            ulong? key = entry.GetUnsigned(1);
            if (entry.FieldCount(1) != 1
                || entry.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !key.HasValue || key.Value > uint.MaxValue) {
                fullyReconstructed = false;
                continue;
            }
            uint normalizedKey = (uint)key.Value;
            if (!seenIdentifiers.Add(normalizedKey)) {
                strings.Remove(normalizedKey);
                fullyReconstructed = false;
                continue;
            }
            IWorkArchiveRecord? wrapper = index.Dereference(entry, 9);
            if (entry.FieldCount(9) != 1
                || wrapper?.MessageType != RichTextWrapperArchive) {
                fullyReconstructed = false;
                continue;
            }
            IWorkWireMessage wrapperMessage = index.Message(wrapper);
            IWorkArchiveRecord? storage = index.Dereference(wrapperMessage, 1);
            if (wrapperMessage.FieldCount(1) != 1
                || storage?.MessageType != TextStorageArchive) {
                fullyReconstructed = false;
                continue;
            }
            IWorkTextContent content = IWorkTextReader.Read(index, storage, projectionBudget);
            if (!content.IsComplete) fullyReconstructed = false;
            if (!content.IsComplete && content.PlainText.Length == 0) continue;
            strings.Add(normalizedKey, content.PlainText);
        }
        return strings;
    }
}
