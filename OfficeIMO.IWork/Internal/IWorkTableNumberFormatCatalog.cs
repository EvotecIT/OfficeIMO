namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves bounded modern numeric formats only when selected by supported numeric cells.</summary>
internal sealed class IWorkTableNumberFormatCatalog {
    private readonly IWorkSourceDocument _source;
    private readonly IWorkWireMessage _store;
    private readonly IWorkArchiveRecord _model;
    private readonly IWorkProjectionBudget _budget;
    private readonly IWorkSourceReferenceIssueCollector _references;
    private readonly Dictionary<uint, (IWorkWireMessage Message, int Position)> _entries = new();
    private readonly Dictionary<uint, IWorkNumberFormat?> _resolved = new();
    private IWorkArchiveRecord? _list;
    private bool _initialized;

    internal IWorkTableNumberFormatCatalog(IWorkSourceDocument source, IWorkWireMessage store,
        IWorkArchiveRecord model, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references) {
        _source = source;
        _store = store;
        _model = model;
        _budget = budget;
        _references = references;
    }

    internal bool FullyReconstructed { get; private set; } = true;

    internal void MarkUnsupportedSelection() => FullyReconstructed = false;

    internal IWorkNumberFormat? Read(uint key) {
        _source.CancellationToken.ThrowIfCancellationRequested();
        if (!_initialized) Initialize();
        if (_resolved.TryGetValue(key, out IWorkNumberFormat? cached)) return cached;
        if (!_entries.TryGetValue(key, out var entry)) {
            FullyReconstructed = false;
            return null;
        }
        string path = IWorkTableCatalogIndex.EntryPath(entry.Position) + "/6";
        IWorkNumberFormat? format = null;
        byte[]? bytes = entry.Message.GetBytes(6);
        if (entry.Message.TotalFieldCount != entry.Message.FieldCount(1)
                + entry.Message.FieldCount(2) + entry.Message.FieldCount(6)
            || entry.Message.FieldCount(2) > 1 || entry.Message.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
            || entry.Message.FieldCount(6) != 1
            || entry.Message.HasUnexpectedWireKind(6, IWorkWireKind.Bytes) || bytes == null) {
            _references.Declarations.Record(_list!, path, entry.Message.FieldCount(6),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        } else {
            try {
                IWorkWireMessage message = entry.Message.ParseNestedMessage(bytes);
                int typeFields = entry.Message.CountNestedFields(bytes, 1, out int totalFields);
                // The supported subset has no currency, date, duration, custom format,
                // scaling or control metadata. Unknown properties cannot be silently ignored.
                bool supportedShape = typeFields == 1 && totalFields == message.FieldCount(1)
                    + message.FieldCount(2) + message.FieldCount(4) + message.FieldCount(5);
                foreach (int field in new[] { 1, 2, 4, 5 })
                    supportedShape &= message.FieldCount(field) <= 1
                        && !message.HasUnexpectedWireKind(field, IWorkWireKind.Varint);
                ulong? type = message.GetUnsigned(1);
                ulong decimals = message.GetUnsigned(2) ?? 0;
                ulong negative = message.GetUnsigned(4) ?? 0;
                ulong grouping = message.GetUnsigned(5) ?? 0;
                if (supportedShape && type is 256 or 258 && (decimals <= 30 || decimals == 253)
                    && negative <= 3 && grouping <= 1) {
                    format = new IWorkNumberFormat(type == 258 ? IWorkNumberFormatKind.Percentage : IWorkNumberFormatKind.Number,
                        decimals == 253 ? null : (int)decimals, grouping == 1, (IWorkNegativeNumberStyle)negative);
                } else {
                    _references.Declarations.Record(_list!, path, 1, IWorkSourceDeclarationIssueKind.InvalidValue);
                }
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                _references.Declarations.Record(_list!, path, 1);
            }
        }
        _resolved.Add(key, format);
        if (format == null) FullyReconstructed = false;
        return format;
    }

    private void Initialize() {
        _initialized = true;
        _list = _references.ReadOne(_model, _store, 22, "4/22");
        if (_store.FieldCount(22) != 1 || _list?.MessageType != 6005) {
            FullyReconstructed = false;
            return;
        }
        IWorkTableCatalogIndex declarations = IWorkTableCatalogIndex.Read(_source, _list, _budget, _references, "number-format");
        FullyReconstructed = declarations.IsComplete;
        if (!declarations.EnvelopeIsComplete) return;
        IWorkWireMessage message = _source.Index.Message(_list);
        if (message.FieldCount(1) != 1 || message.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || message.GetUnsigned(1) != 2) {
            _references.Declarations.Record(_list, "1", message.FieldCount(1),
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            FullyReconstructed = false;
            return;
        }
        foreach (var entry in declarations.Entries) {
            _source.CancellationToken.ThrowIfCancellationRequested();
            if (declarations.CanResolveKey(entry.Key)) _entries.Add(entry.Key, (entry.Message, entry.Position));
        }
    }
}
