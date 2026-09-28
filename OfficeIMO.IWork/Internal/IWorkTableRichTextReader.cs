namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves rich-text cell references through the table catalog and text-storage graph.</summary>
internal static class IWorkTableRichTextReader {
    private const uint DataListArchive = 6005;
    private const uint RichTextWrapperArchive = 6218;
    private const uint TextStorageArchive = 2001;

    internal static IReadOnlyDictionary<uint, IWorkTextContent> Read(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkProjectionBudget projectionBudget,
        IWorkReadOptions options, int maximumEntries, out bool fullyReconstructed) =>
        Read(index, store, projectionBudget, options, maximumEntries,
            out fullyReconstructed, out _);

    internal static IReadOnlyDictionary<uint, IWorkTextContent> Read(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkProjectionBudget projectionBudget,
        IWorkReadOptions options, int maximumEntries, out bool fullyReconstructed,
        out bool catalogStructureComplete) {
        var strings = new Dictionary<uint, IWorkTextContent>();
        var seenIdentifiers = new HashSet<uint>();
        var recordMessages = new Dictionary<ulong, IWorkWireMessage?>();
        var storageContents = new Dictionary<ulong, IWorkTextContent?>();
        fullyReconstructed = true;
        catalogStructureComplete = true;
        if (!store.HasField(17)) return strings;
        IWorkArchiveRecord? list = index.Dereference(store, 17);
        if (store.FieldCount(17) != 1 || list?.MessageType != DataListArchive) {
            fullyReconstructed = false;
            catalogStructureComplete = false;
            return strings;
        }
        if (!IWorkNumbersReader.TryGetCatalogEntryCount(list, maximumEntries, options,
                "rich-text", out int entryCount)) {
            fullyReconstructed = false;
            catalogStructureComplete = false;
            return strings;
        }
        projectionBudget.AddTableCatalogEntries(entryCount);
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            index.Message(list), 3, out bool malformedEntries);
        if (malformedEntries) {
            fullyReconstructed = false;
            catalogStructureComplete = false;
        }
        foreach (IWorkWireMessage entry in entries) {
            ulong? key = entry.GetUnsigned(1);
            if (entry.FieldCount(1) != 1
                || entry.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !key.HasValue || key.Value > uint.MaxValue) {
                fullyReconstructed = false;
                catalogStructureComplete = false;
                continue;
            }
            uint normalizedKey = (uint)key.Value;
            if (!seenIdentifiers.Add(normalizedKey)) {
                strings.Remove(normalizedKey);
                fullyReconstructed = false;
                catalogStructureComplete = false;
                continue;
            }
            IWorkArchiveRecord? wrapper = index.Dereference(entry, 9);
            if (entry.FieldCount(9) != 1
                || wrapper?.MessageType != RichTextWrapperArchive) {
                fullyReconstructed = false;
                continue;
            }
            if (!TryReadRecord(index, wrapper, options, recordMessages,
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
            if (!storageContents.TryGetValue(storage.Identifier, out IWorkTextContent? content)) {
                content = TryReadRecord(index, storage, options, recordMessages,
                        out IWorkWireMessage? storageMessage) && storageMessage != null
                    ? IWorkTextReader.Read(index, storage, projectionBudget,
                        tolerateStyleDepth: true)
                    : null;
                storageContents.Add(storage.Identifier, content);
            } else if (content != null) {
                projectionBudget.AddTextContentUse(content, includeCharacters: true);
            }
            if (content == null) {
                fullyReconstructed = false;
                continue;
            }
            if (!content.IsTextComplete) fullyReconstructed = false;
            if (!content.IsTextComplete && content.PlainText.Length == 0) continue;
            strings.Add(normalizedKey, content);
        }
        return strings;
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
