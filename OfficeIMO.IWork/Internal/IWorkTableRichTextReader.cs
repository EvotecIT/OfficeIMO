namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves rich-text cell references through the table catalog and text-storage graph.</summary>
internal static class IWorkTableRichTextReader {
    private const uint DataListArchive = 6005;
    private const uint RichTextWrapperArchive = 6218;
    private const uint TextStorageArchive = 2001;

    internal static IReadOnlyDictionary<uint, (string Text, bool IsComplete)> Read(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkProjectionBudget projectionBudget,
        IWorkReadOptions options, int maximumEntries, out bool fullyReconstructed) {
        var strings = new Dictionary<uint, (string Text, bool IsComplete)>();
        var seenIdentifiers = new HashSet<uint>();
        var wrapperMessages = new Dictionary<ulong, IWorkWireMessage?>();
        var storageTexts = new Dictionary<ulong, (string Text, bool IsComplete)?>();
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
            if (!TryReadRecord(index, wrapper, options, wrapperMessages,
                    out IWorkWireMessage? wrapperMessage)
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
            if (!storageTexts.TryGetValue(storage.Identifier, out var cached)) {
                cached = TryReadRecord(index, storage, options, wrapperMessages,
                        out IWorkWireMessage? storageMessage) && storageMessage != null
                    ? ReadStorage(storageMessage, projectionBudget)
                    : null;
                storageTexts.Add(storage.Identifier, cached);
            }
            if (!cached.HasValue) {
                fullyReconstructed = false;
                continue;
            }
            (string text, bool textComplete) = cached.Value;
            if (!textComplete) fullyReconstructed = false;
            if (!textComplete && text.Length == 0) continue;
            strings.Add(normalizedKey, (text, textComplete));
        }
        return strings;
    }

    private static (string Text, bool IsComplete) ReadStorage(IWorkWireMessage message,
        IWorkProjectionBudget projectionBudget) {
        string text = IWorkTextReader.ReadPlainText(message, projectionBudget,
            out bool complete);
        return (text, complete);
    }

    private static bool TryReadRecord(IWorkObjectIndex index, IWorkArchiveRecord record,
        IWorkReadOptions options, Dictionary<ulong, IWorkWireMessage?> cache,
        out IWorkWireMessage? message) {
        if (cache.TryGetValue(record.Identifier, out message)) return message != null;
        message = null;
        try {
            // Count first so configured field limits remain fatal even when the record is malformed.
            IWorkProtobuf.CountFields(record.Payload, 1, options.MaximumProtobufFieldCount);
            message = index.Message(record);
            cache.Add(record.Identifier, message);
            return true;
        } catch (InvalidDataException exception)
            when (!IWorkProtobuf.IsFieldLimitException(exception)) {
            cache.Add(record.Identifier, null);
            return false;
        }
    }
}
