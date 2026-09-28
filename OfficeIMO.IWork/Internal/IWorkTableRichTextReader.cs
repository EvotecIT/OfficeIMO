namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves rich-text cell references through the table catalog and text-storage graph.</summary>
internal static class IWorkTableRichTextReader {
    private const uint DataListArchive = 6005;
    private const uint RichTextWrapperArchive = 6218;
    private const uint TextStorageArchive = 2001;

    internal static IReadOnlyDictionary<uint, IWorkTextContent> Read(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkProjectionBudget projectionBudget,
        IWorkReadOptions options, int maximumEntries, out bool fullyReconstructed) {
        var strings = new Dictionary<uint, IWorkTextContent>();
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
            if (!TryReadRecord(index, wrapper, options, out IWorkWireMessage? wrapperMessage)
                || wrapperMessage == null) {
                fullyReconstructed = false;
                continue;
            }
            IWorkArchiveRecord? storage = index.Dereference(wrapperMessage, 1);
            if (wrapperMessage.FieldCount(1) != 1
                || storage?.MessageType != TextStorageArchive) {
                fullyReconstructed = false;
                continue;
            }
            if (!TryReadRecord(index, storage, options, out IWorkWireMessage? storageMessage)
                || storageMessage == null) {
                fullyReconstructed = false;
                continue;
            }
            IWorkTextContent content = IWorkTextReader.Read(index, storage, projectionBudget,
                tolerateStyleDepth: true);
            if (!content.IsTextComplete) fullyReconstructed = false;
            if (!content.IsTextComplete && content.PlainText.Length == 0) continue;
            strings.Add(normalizedKey, content);
        }
        return strings;
    }

    private static bool TryReadRecord(IWorkObjectIndex index, IWorkArchiveRecord record,
        IWorkReadOptions options, out IWorkWireMessage? message) {
        message = null;
        try {
            // Count first so configured field limits remain fatal even when the record is malformed.
            IWorkProtobuf.CountFields(record.Payload, 1, options.MaximumProtobufFieldCount);
            message = index.Message(record);
            return true;
        } catch (InvalidDataException exception)
            when (!IWorkProtobuf.IsFieldLimitException(exception)) {
            return false;
        }
    }
}
