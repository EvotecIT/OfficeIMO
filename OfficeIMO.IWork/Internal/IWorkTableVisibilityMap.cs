using System.Globalization;

namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves selected base user-hidden UUIDs through the canonical table position map.</summary>
internal static class IWorkTableVisibilityMap {
    internal static bool TryResolve(IWorkSourceDocument source, IWorkArchiveRecord model, IWorkWireMessage modelMessage,
        int dimensionCount, bool columns, IReadOnlyList<(IWorkWireMessage Message, string Path)> states,
        IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references, out IReadOnlyList<int> hidden) {
        hidden = Array.Empty<int>();
        if (!modelMessage.HasField(46)) return false;
        budget.AddTableDimensionEntries(modelMessage.FieldCount(46));
        IWorkArchiveRecord? record = references.ReadOne(model, modelMessage, 46, "46", static type => type == 6267);
        if (record == null) return false;
        if (record.MessageType != 6267) {
            references.Declarations.Record(record, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            return false;
        }
        IWorkWireMessage map;
        try { map = source.Index.Message(record); }
        catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(record, "$", null);
            return false;
        }
        int uuidField = columns ? 1 : 4;
        int indexField = uuidField + 1;
        int inverseField = uuidField + 2;
        budget.AddTableDimensionEntries(map.FieldCount(uuidField));
        if (map.FieldCount(uuidField) != dimensionCount) return Invalid(uuidField);
        IReadOnlyList<ulong>? indexes = ReadIndexes(indexField);
        if (indexes == null) return false;
        if (indexes.Count != dimensionCount || indexes.Any(index => index >= (ulong)dimensionCount)
            || indexes.Distinct().Count() != dimensionCount) return Invalid(indexField);
        if (map.HasField(inverseField)) {
            IReadOnlyList<ulong>? inverse = ReadIndexes(inverseField);
            if (inverse == null) return false;
            if (inverse.Count != dimensionCount) return Invalid(inverseField);
            for (int position = 0; position < indexes.Count; position++) {
                source.CancellationToken.ThrowIfCancellationRequested();
                if (inverse[(int)indexes[position]] != (ulong)position) return Invalid(inverseField);
            }
        }
        // UUID words come from the source. Ordered keys prevent adversarial
        // tuple-hash collisions from making declaration mapping quadratic.
        var positions = new SortedDictionary<(ulong Lower, ulong Upper), int>();
        int ordinal = 0;
        foreach (IWorkWireValue value in map.EnumerateValues(uuidField)) {
            source.CancellationToken.ThrowIfCancellationRequested();
            IWorkWireMessage? uuid = null;
            if (value.Kind == IWorkWireKind.Bytes && value.Bytes != null) {
                try { uuid = map.ParseNestedMessage(value.Bytes); }
                catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) { }
            }
            if (!TryUuid(uuid, out var key) || positions.ContainsKey(key))
                return Invalid(uuidField);
            positions.Add(key, (int)indexes[ordinal] + 1);
            ordinal++;
        }
        var seen = new SortedSet<(ulong Lower, ulong Upper)>();
        var recovered = new List<int>();
        foreach (var state in states) {
            source.CancellationToken.ThrowIfCancellationRequested();
            IWorkWireMessage? uuid = IWorkObjectIndex.TryGetMessage(state.Message, 1);
            bool validFlag = !state.Message.HasField(2) || state.Message.FieldCount(2) == 1
                && !state.Message.HasUnexpectedWireKind(2, IWorkWireKind.Varint) && state.Message.GetUnsigned(2) <= 1;
            if (!validFlag || !TryUuid(uuid, out var key) || !seen.Add(key) || !positions.TryGetValue(key, out int position)) {
                references.Declarations.Record(model, state.Path + "/1", state.Message.FieldCount(1),
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                return false;
            }
            if (state.Message.GetUnsigned(2) == 1) recovered.Add(position);
        }
        hidden = recovered.OrderBy(position => position).ToArray();
        return true;

        bool Invalid(int field) {
            references.Declarations.Record(record, field.ToString(CultureInfo.InvariantCulture), map.FieldCount(field),
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            return false;
        }
        IReadOnlyList<ulong>? ReadIndexes(int field) {
            if (map.HasUnexpectedWireKind(field, IWorkWireKind.Varint, IWorkWireKind.Bytes)) { Invalid(field); return null; }
            try {
                IReadOnlyList<ulong> values = map.GetRepeatedUnsigned(field, packed: true, budget.RemainingTableDimensionEntries);
                budget.AddTableDimensionEntries(values.Count);
                return values;
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                references.Declarations.Record(record, field.ToString(CultureInfo.InvariantCulture), map.FieldCount(field));
                return null;
            }
        }
    }

    private static bool TryUuid(IWorkWireMessage? message, out (ulong Lower, ulong Upper) uuid) {
        uuid = default;
        if (message == null || message.TotalFieldCount != 2 || message.FieldCount(1) != 1 || message.FieldCount(2) != 1
            || message.HasUnexpectedWireKind(1, IWorkWireKind.Varint) || message.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
            || message.GetUnsigned(1) is not ulong lower || message.GetUnsigned(2) is not ulong upper) return false;
        uuid = (lower, upper);
        return true;
    }
}
