namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves only selected, independently qualified root comments through the shared table catalog.</summary>
internal sealed class IWorkTableCommentCatalog(IWorkSourceDocument source, IWorkWireMessage store,
    IWorkArchiveRecord model, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references) {
    private readonly Dictionary<uint, (IWorkWireMessage Message, int Position)> _entries = new();
    private readonly Dictionary<ulong, IWorkCellComment?> _comments = new();
    private readonly Dictionary<ulong, IWorkWireMessage?> _messages = new();
    private IWorkArchiveRecord? _list;
    private bool _initialized;

    internal IWorkCellComment? Read(uint key) {
        source.CancellationToken.ThrowIfCancellationRequested();
        budget.AddTextItem();
        Initialize();
        if (!_entries.TryGetValue(key, out var entry)) return null;
        string path = IWorkTableCatalogIndex.EntryPath(entry.Position) + "/10";
        IWorkArchiveRecord? record = references.ReadOne(_list!, entry.Message, 10, path);
        if (entry.Message.FieldCount(10) != 1 || record?.MessageType != 3056) {
            references.Declarations.Record(_list!, path, entry.Message.FieldCount(10),
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            return null;
        }
        if (_comments.TryGetValue(record.Identifier, out IWorkCellComment? cached)) {
            if (cached != null) { budget.AddTextCharacters(cached.Text.Length); budget.AddTextCharacters(cached.Author.Length); }
            return cached;
        }
        IWorkCellComment? comment = ReadComment(record);
        _comments.Add(record.Identifier, comment);
        return comment;
    }

    private void Initialize() {
        if (_initialized) return;
        _initialized = true;
        _list = references.ReadOne(model, store, 19, "4/19");
        if (store.FieldCount(19) != 1 || _list?.MessageType != 6005) {
            references.Declarations.Record(model, "4/19", store.FieldCount(19),
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            return;
        }
        IWorkTableCatalogIndex catalog = IWorkTableCatalogIndex.Read(source, _list, budget, references, "comment");
        IWorkWireMessage? message = ReadMessage(_list);
        if (message == null || !catalog.EnvelopeIsComplete) return;
        if (message.FieldCount(1) != 1 || message.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || message.GetUnsigned(1) != 10) {
            references.Declarations.Record(_list, "1", message.FieldCount(1), IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            return;
        }
        foreach (var entry in catalog.Entries) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (!catalog.CanResolveKey(entry.Key)) continue;
            if (entry.Message.TotalFieldCount != entry.Message.FieldCount(1) + entry.Message.FieldCount(2)
                + entry.Message.FieldCount(10)) {
                references.Declarations.Record(_list, IWorkTableCatalogIndex.EntryPath(entry.Position), 1,
                    IWorkSourceDeclarationIssueKind.UnsupportedField);
                continue;
            }
            _entries.Add(entry.Key, (entry.Message, entry.Position));
        }
    }

    private IWorkCellComment? ReadComment(IWorkArchiveRecord record) {
        IWorkWireMessage? message = ReadMessage(record);
        if (message == null) return null;
        // Replies need their own native qualification. Do not follow an unqualified graph or flatten it.
        if (message.HasField(4)) {
            references.Declarations.Record(record, "4", message.FieldCount(4), IWorkSourceDeclarationIssueKind.UnsupportedField);
            return null;
        }
        if (message.TotalFieldCount != message.FieldCount(1) + message.FieldCount(2)
            + message.FieldCount(3) + message.FieldCount(5)) {
            references.Declarations.Record(record, "$", null, IWorkSourceDeclarationIssueKind.UnsupportedField);
            return null;
        }
        string? text = ReadText(record, message, 1);
        if (text == null) return null;
        IWorkWireMessage? date = IWorkObjectIndex.TryGetMessage(message, 2);
        double? seconds = date?.GetDouble(1);
        if (message.FieldCount(2) != 1 || message.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
            || date?.TotalFieldCount != 1 || date.HasUnexpectedWireKind(1, IWorkWireKind.Fixed64)
            || seconds == null || double.IsNaN(seconds.Value) || double.IsInfinity(seconds.Value)) {
            references.Declarations.Record(record, "2", message.FieldCount(2), IWorkSourceDeclarationIssueKind.InvalidValue);
            return null;
        }
        DateTime timestamp;
        try { timestamp = new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(seconds.Value); }
        catch (ArgumentOutOfRangeException) {
            references.Declarations.Record(record, "2", message.FieldCount(2), IWorkSourceDeclarationIssueKind.InvalidValue);
            return null;
        }
        // Native UUIDs are metadata, not destination comment IDs. Validate their qualified shape without inventing an ID mapping.
        if (message.HasField(5)) {
            IWorkWireMessage? uuid = IWorkObjectIndex.TryGetMessage(message, 5);
            if (message.FieldCount(5) != 1 || message.HasUnexpectedWireKind(5, IWorkWireKind.Bytes)
                || uuid?.TotalFieldCount != 2 || uuid.FieldCount(1) != 1 || uuid.FieldCount(2) != 1
                || uuid.HasUnexpectedWireKind(1, IWorkWireKind.Varint) || uuid.HasUnexpectedWireKind(2, IWorkWireKind.Varint)) {
                references.Declarations.Record(record, "5", message.FieldCount(5), IWorkSourceDeclarationIssueKind.InvalidValue);
                return null;
            }
        }
        IWorkArchiveRecord? author = references.ReadOne(record, message, 3);
        if (message.FieldCount(3) != 1 || author?.MessageType != 212) {
            references.Declarations.Record(record, "3", message.FieldCount(3), IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            return null;
        }
        IWorkWireMessage? authorMessage = ReadMessage(author);
        if (authorMessage == null) return null;
        // The qualified author fields beyond the display name are identity/color metadata.
        if (authorMessage.TotalFieldCount != Enumerable.Range(1, 5).Sum(authorMessage.FieldCount)) {
            references.Declarations.Record(author, "$", null, IWorkSourceDeclarationIssueKind.UnsupportedField);
            return null;
        }
        string? name = ReadText(author, authorMessage, 1);
        return name == null ? null : new IWorkCellComment(text, name, timestamp,
            new IWorkObjectIdentity(record), new IWorkObjectIdentity(author));
    }

    private string? ReadText(IWorkArchiveRecord record, IWorkWireMessage message, int field) {
        string? text = message.GetString(field, out bool complete);
        if (!complete || text == null) {
            references.Declarations.Record(record, field.ToString(System.Globalization.CultureInfo.InvariantCulture),
                message.FieldCount(field), IWorkSourceDeclarationIssueKind.InvalidValue);
            return null;
        }
        budget.AddTextCharacters(text.Length);
        return text;
    }

    private IWorkWireMessage? ReadMessage(IWorkArchiveRecord record) {
        if (_messages.TryGetValue(record.Identifier, out IWorkWireMessage? message)) return message;
        try {
            IWorkProtobuf.CountFields(record.Payload, 1, source.Options.MaximumProtobufFieldCount);
            message = source.Index.Message(record);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(record, "$", null);
        }
        _messages.Add(record.Identifier, message);
        return message;
    }
}
