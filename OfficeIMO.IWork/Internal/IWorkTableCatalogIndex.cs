using System.Globalization;

namespace OfficeIMO.IWork.Internal;

/// <summary>Indexes bounded physical catalog declarations before any consumer trusts keys or decodes values.</summary>
internal sealed class IWorkTableCatalogIndex {
    private readonly List<(uint Key, IWorkWireMessage Message, int Position)> _entries = new();
    private readonly HashSet<uint> _ambiguousKeys = new();
    private bool _unknownKeys;

    internal bool EnvelopeIsComplete { get; private set; }
    internal bool IsComplete { get; private set; }
    internal IReadOnlyList<(uint Key, IWorkWireMessage Message, int Position)> Entries => _entries;

    /// <summary>An unreadable key can conceal a duplicate; known duplicate keys remain unresolved even after later healthy entries.</summary>
    internal bool CanResolveKey(uint key) => !_unknownKeys && !_ambiguousKeys.Contains(key);

    internal static IWorkTableCatalogIndex Read(IWorkSourceDocument source, IWorkArchiveRecord list,
        IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references, string catalogName) {
        var result = new IWorkTableCatalogIndex();
        int declaredEntries;
        int totalFields;
        int identifierFields;
        int metadataFields;
        try {
            declaredEntries = IWorkProtobuf.CountFields(list.Payload, 3,
                source.Options.MaximumProtobufFieldCount, out totalFields);
            identifierFields = IWorkProtobuf.CountFields(list.Payload, 1, source.Options.MaximumProtobufFieldCount);
            metadataFields = IWorkProtobuf.CountFields(list.Payload, 2, source.Options.MaximumProtobufFieldCount);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(list, "$", null);
            return result;
        }
        if (identifierFields > 1 || metadataFields > 1
            || totalFields - declaredEntries != identifierFields + metadataFields) {
            references.Declarations.Record(list, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            return result;
        }
        if (declaredEntries > budget.RemainingTableCatalogEntries) {
            throw new InvalidDataException($"An iWork {catalogName} catalog exceeds the remaining table-catalog limit of {budget.RemainingTableCatalogEntries}.");
        }
        budget.AddTableCatalogEntries(declaredEntries);
        IWorkWireMessage message = source.Index.Message(list);
        result.EnvelopeIsComplete = result.IsComplete = true;
        var seen = new HashSet<uint>();
        int position = 0;
        foreach (IWorkWireValue value in message.EnumerateValues(3)) {
            source.CancellationToken.ThrowIfCancellationRequested();
            position++;
            IWorkWireMessage? entry = null;
            if (value.Kind == IWorkWireKind.Bytes && value.Bytes != null) {
                try { entry = message.ParseNestedMessage(value.Bytes); }
                catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) { }
            }
            if (entry == null) {
                result.IsComplete = false;
                result._unknownKeys = true;
                references.Declarations.Record(list, EntryPath(position), 1);
                continue;
            }
            ulong? key = entry.GetUnsigned(1);
            if (entry.FieldCount(1) != 1 || entry.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !key.HasValue || key.Value > uint.MaxValue) {
                result.IsComplete = false;
                result._unknownKeys = true;
                references.Declarations.Record(list, EntryPath(position) + "/1", entry.FieldCount(1),
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                continue;
            }
            uint normalized = (uint)key.Value;
            if (!seen.Add(normalized)) {
                result.IsComplete = false;
                result._ambiguousKeys.Add(normalized);
                references.Declarations.Record(list, EntryPath(position) + "/1", 1,
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            }
            // Keep readable occurrences so eager value readers still charge all decoding work.
            // Lazy rich-text consumers consult CanResolveKey before traversing any value.
            result._entries.Add((normalized, entry, position));
        }
        return result;
    }

    internal static string EntryPath(int position) => "3[" + position.ToString(CultureInfo.InvariantCulture) + "]";
}
