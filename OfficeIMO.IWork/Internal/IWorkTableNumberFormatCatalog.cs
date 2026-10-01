namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves bounded modern numeric formats only when selected by supported numeric cells.</summary>
internal sealed class IWorkTableNumberFormatCatalog {
    private readonly IWorkSourceDocument _source;
    private readonly IWorkWireMessage _store;
    private readonly IWorkArchiveRecord _model;
    private readonly IWorkProjectionBudget _budget;
    private readonly IWorkSourceReferenceIssueCollector _references;
    private readonly Dictionary<uint, (IWorkWireMessage Message, int Position)> _entries = new();
    private readonly Dictionary<(uint Key, bool Currency), IWorkNumberFormat?> _resolved = new();
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

    internal IWorkNumberFormat? Read(uint key, bool currency = false) {
        _source.CancellationToken.ThrowIfCancellationRequested();
        if (!_initialized) Initialize();
        if (_resolved.TryGetValue((key, currency), out IWorkNumberFormat? cached)) return cached;
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
                // The supported subset has no date, duration, custom format,
                // scaling or control metadata. Unknown properties cannot be silently ignored.
                bool supportedShape = typeFields == 1 && totalFields == message.FieldCount(1)
                    + message.FieldCount(2) + message.FieldCount(4) + message.FieldCount(5)
                    + (currency ? message.FieldCount(3) + message.FieldCount(6) : 0);
                foreach (int field in currency ? new[] { 1, 2, 4, 5, 6 } : new[] { 1, 2, 4, 5 })
                    supportedShape &= message.FieldCount(field) <= 1
                        && !message.HasUnexpectedWireKind(field, IWorkWireKind.Varint);
                ulong? type = message.GetUnsigned(1);
                ulong decimals = message.GetUnsigned(2) ?? 0;
                ulong negative = message.GetUnsigned(4) ?? 0;
                ulong grouping = message.GetUnsigned(5) ?? 0;
                ulong accounting = currency ? message.GetUnsigned(6) ?? 0 : 0;
                string? currencyCode = null;
                if (currency) {
                    byte[]? code = message.GetBytes(3);
                    supportedShape &= message.FieldCount(3) == 1
                        && !message.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
                        && code is { Length: 3 } && code.All(value => value is >= (byte)'A' and <= (byte)'Z');
                    if (supportedShape) currencyCode = System.Text.Encoding.ASCII.GetString(code!);
                }
                if (supportedShape && (currency ? type == 257 : type is 256 or 258 or 259)
                    && (decimals <= 30 || decimals == 253) && negative <= 3 && grouping <= 1
                    && (type != 259 || negative == 0 && grouping == 0)
                    && accounting <= 1 && (accounting == 0 || negative == 0)) {
                    format = new IWorkNumberFormat(currency ? IWorkNumberFormatKind.Currency
                            : type == 259 ? IWorkNumberFormatKind.Scientific
                            : type == 258 ? IWorkNumberFormatKind.Percentage : IWorkNumberFormatKind.Number,
                        decimals == 253 ? null : (int)decimals, grouping == 1, (IWorkNegativeNumberStyle)negative,
                        currencyCode, accounting == 1);
                } else {
                    _references.Declarations.Record(_list!, path, 1, IWorkSourceDeclarationIssueKind.InvalidValue);
                }
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                _references.Declarations.Record(_list!, path, 1);
            }
        }
        // Projection-budget failures are fatal, not malformed-message recovery.
        if (format?.CurrencyCode is string retainedCode) {
            _budget.AddTextCharacters(retainedCode.Length);
            _budget.AddTextItem();
        }
        _resolved.Add((key, currency), format);
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
