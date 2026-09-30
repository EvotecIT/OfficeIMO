namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static IReadOnlyList<AttributeBoundary> ReadObjectTable(IWorkWireMessage storage,
        int field, int textLength, IWorkArchiveRecord owner, IWorkProjectionBudget projectionBudget,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        if (!storage.HasField(field)) return Array.Empty<AttributeBoundary>();
        string tablePath = field.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (storage.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)) {
            complete = false;
            references.Declarations.Record(owner, tablePath, storage.FieldCount(field));
            return Array.Empty<AttributeBoundary>();
        }
        byte[] tableBytes = storage.GetBytes(field)!;
        int boundaryCount;
        int totalTableFieldCount;
        try {
            boundaryCount = storage.CountNestedFields(tableBytes, 1,
                out totalTableFieldCount);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            complete = false;
            references.Declarations.Record(owner, tablePath, storage.FieldCount(field));
            return Array.Empty<AttributeBoundary>();
        }
        projectionBudget.AddTextBoundaries(boundaryCount);
        if (storage.FieldCount(field) != 1 || totalTableFieldCount != boundaryCount) {
            complete = false;
            references.Declarations.Record(owner, tablePath, storage.FieldCount(field),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            return Array.Empty<AttributeBoundary>();
        }
        IWorkWireMessage table;
        try {
            table = storage.ParseNestedMessage(tableBytes);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            complete = false;
            references.Declarations.Record(owner, tablePath, storage.FieldCount(field));
            return Array.Empty<AttributeBoundary>();
        }
        var result = new List<AttributeBoundary>();
        if (table.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)) complete = false;
        int entryIndex = 0;
        foreach (IWorkWireValue value in table.EnumerateValues(1)) {
            entryIndex++;
            if (value.Kind != IWorkWireKind.Bytes || value.Bytes == null) {
                references.Declarations.Record(owner, EntryPath(entryIndex), 1);
                continue;
            }
            IWorkWireMessage entry;
            try {
                entry = table.ParseNestedMessage(value.Bytes);
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                complete = false;
                references.Declarations.Record(owner, EntryPath(entryIndex), 1);
                continue;
            }
            ulong? rawIndex = entry.GetUnsigned(1);
            if (entry.FieldCount(1) != 1
                || entry.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !rawIndex.HasValue || rawIndex.Value > int.MaxValue
                || rawIndex.Value > (ulong)textLength) {
                complete = false;
                references.Declarations.Record(owner, EntryPath(entryIndex), 1,
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                continue;
            }
            bool hasObject = entry.HasField(2);
            if (hasObject) {
                // Every decoded, valid-offset attribute entry is assessed by this reader,
                // including declarations at the end of the text. Preserve physical indexes.
                references.ReadOne(owner, entry, 2,
                    field.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/1["
                    + entryIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]/2");
            }
            bool malformedReference = false;
            IWorkWireMessage? reference = hasObject
                ? IWorkObjectIndex.TryGetMessage(entry, 2, out malformedReference)
                : null;
            if (hasObject && (entry.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                    || malformedReference || reference?.FieldCount(1) != 1
                    || reference?.GetUnsigned(1) == null
                    || reference.HasUnexpectedWireKind(1, IWorkWireKind.Varint))) {
                complete = false;
                continue;
            }
            result.Add(new AttributeBoundary((int)rawIndex.Value,
                reference?.GetUnsigned(1), hasObject, entryIndex));
        }
        AttributeBoundary[] ordered = result.OrderBy(boundary => boundary.Index).ToArray();
        for (int index = 1; index < ordered.Length; index++) {
            if (ordered[index - 1].Index == ordered[index].Index) {
                complete = false;
                references.Declarations.Record(owner, EntryPath(ordered[index].Position), 1,
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            }
        }
        ulong? carried = null;
        foreach (AttributeBoundary boundary in ordered) {
            if (boundary.HasObject) carried = boundary.Identifier;
            boundary.CarriedIdentifier = carried;
        }
        return ordered;

        string EntryPath(int position) => tablePath + "/1["
            + position.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
    }

    private static ulong? ObjectAt(IReadOnlyList<AttributeBoundary> boundaries, int offset,
        bool carryMissing) {
        int upper = UpperBound(boundaries, offset);
        if (upper == 0) return null;
        AttributeBoundary boundary = boundaries[upper - 1];
        return carryMissing ? boundary.CarriedIdentifier
            : boundary.HasObject ? boundary.Identifier : null;
    }

    private static void AddBoundaries(SortedSet<int> destination,
        IReadOnlyList<AttributeBoundary> source, int start, int end) {
        int index = UpperBound(source, start);
        while (index < source.Count && source[index].Index < end) {
            destination.Add(source[index].Index);
            index++;
        }
    }

    private static int UpperBound(IReadOnlyList<AttributeBoundary> boundaries, int offset) {
        int low = 0;
        int high = boundaries.Count;
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (boundaries[middle].Index <= offset) low = middle + 1;
            else high = middle;
        }
        return low;
    }

    private sealed class AttributeBoundary {
        internal AttributeBoundary(int index, ulong? identifier, bool hasObject, int position) {
            Index = index;
            Identifier = identifier;
            HasObject = hasObject;
            Position = position;
        }
        internal int Position { get; }
        internal int Index { get; }
        internal ulong? Identifier { get; }
        internal bool HasObject { get; }
        internal ulong? CarriedIdentifier { get; set; }
    }

}
