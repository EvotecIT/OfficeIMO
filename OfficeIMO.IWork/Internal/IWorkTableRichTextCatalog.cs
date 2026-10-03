namespace OfficeIMO.IWork.Internal;

/// <summary>Indexes bounded rich-text declarations and resolves only entries selected by decoded cells.</summary>
internal sealed class IWorkTableRichTextCatalog {
    private const uint RichTextWrapperArchive = 6218;
    private const uint TextStorageArchive = 2001;
    private readonly IWorkObjectIndex _index;
    private readonly IWorkProjectionBudget _budget;
    private readonly IWorkReadOptions _options;
    private readonly IWorkSourceReferenceIssueCollector _references;
    private readonly Dictionary<uint, (IWorkWireMessage Message, int Position)> _entries = new();
    private readonly HashSet<uint> _attempted = new();
    private readonly Dictionary<ulong, IWorkWireMessage?> _recordMessages = new();
    private readonly Dictionary<ulong, IWorkTextContent?> _storageContents = new();
    private readonly Dictionary<uint, IWorkTextContent> _materialized = new();
    private readonly Dictionary<uint, IWorkObjectIdentity> _omitted = new();
    private IWorkArchiveRecord? _list;

    private IWorkTableRichTextCatalog(IWorkObjectIndex index, IWorkProjectionBudget budget,
        IWorkReadOptions options, IWorkSourceReferenceIssueCollector references) {
        _index = index;
        _budget = budget;
        _options = options;
        _references = references;
    }

    internal bool FullyReconstructed { get; private set; } = true;
    internal bool StructureComplete { get; private set; } = true;
    internal IReadOnlyDictionary<uint, IWorkTextContent> Materialized => _materialized;
    internal IReadOnlyDictionary<uint, IWorkObjectIdentity> OmittedStorages => _omitted;

    internal static IWorkTableRichTextCatalog Create(IWorkSourceDocument source, IWorkWireMessage store,
        IWorkArchiveRecord model, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references) {
        var catalog = new IWorkTableRichTextCatalog(source.Index, budget, source.Options, references);
        if (!store.HasField(17)) return catalog;
        IWorkArchiveRecord? list = references.ReadOne(model, store, 17, "4/17", IWorkTableCatalogIndex.IsDataListType);
        if (store.FieldCount(17) != 1 || list == null || !IWorkTableCatalogIndex.IsDataListType(list.MessageType)) {
            catalog.FullyReconstructed = catalog.StructureComplete = false;
            return catalog;
        }
        catalog._list = list;
        IWorkTableCatalogIndex declarations = IWorkTableCatalogIndex.Read(source, list, budget, references, "rich-text", 8);
        catalog.FullyReconstructed = catalog.StructureComplete = declarations.IsComplete;
        foreach (var entry in declarations.Entries) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (declarations.CanResolveKey(entry.Key)) catalog._entries.Add(entry.Key, (entry.Message, entry.Position));
        }
        return catalog;
    }

    internal bool TryRead(uint key, out IWorkTextContent? content) {
        if (_materialized.TryGetValue(key, out content)) return true;
        content = null;
        if (!_entries.TryGetValue(key, out var entry)) {
            FullyReconstructed = false;
            return false;
        }
        // Cache failed attempts only for declared keys, bounded by the catalog entry limit.
        if (!_attempted.Add(key)) return false;
        IWorkArchiveRecord? wrapper = _references.ReadOne(_list!, entry.Message, 9,
                "3[" + entry.Position.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]/9", static type => type == RichTextWrapperArchive);
        if (entry.Message.FieldCount(9) != 1 || wrapper?.MessageType != RichTextWrapperArchive
            || !TryReadRecord(_index, wrapper, _options, _recordMessages, out IWorkWireMessage? wrapperMessage)
            || wrapperMessage == null) {
            FullyReconstructed = false;
            return false;
        }
        IWorkArchiveRecord? storage = _references.ReadOne(wrapper, wrapperMessage, 1, allowedType: static type => type == TextStorageArchive);
        if (wrapperMessage.FieldCount(1) != 1 || storage?.MessageType != TextStorageArchive) {
            FullyReconstructed = false;
            return false;
        }
        if (!_storageContents.TryGetValue(storage.Identifier, out content)) {
            content = TryReadRecord(_index, storage, _options, _recordMessages,
                    out IWorkWireMessage? storageMessage) && storageMessage != null
                ? IWorkTextReader.Read(_index, storage, _budget, _references, tolerateStyleDepth: true)
                : null;
            _storageContents.Add(storage.Identifier, content);
        }
        if (content == null || !content.IsTextComplete && content.PlainText.Length == 0) {
            _omitted.Add(key, new IWorkObjectIdentity(storage));
            FullyReconstructed = false;
            content = null;
            return false;
        }
        if (!content.IsTextComplete) FullyReconstructed = false;
        _materialized.Add(key, content);
        return true;
    }

    private bool TryReadRecord(IWorkObjectIndex index, IWorkArchiveRecord record,
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
            when (!IWorkProtobuf.IsLimitException(exception)) {
            _references.Declarations.Record(record, "$", null);
            cache.Add(record.Identifier, null);
            return false;
        }
    }
}
