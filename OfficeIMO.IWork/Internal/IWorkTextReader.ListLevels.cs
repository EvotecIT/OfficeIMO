namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static IReadOnlyList<AttributeBoundary> ReadListLevelTable(IWorkWireMessage storage,
        int textLength, IWorkArchiveRecord owner, IWorkProjectionBudget projectionBudget,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        const int field = 6;
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
            if (entry.TotalFieldCount != entry.FieldCount(1) + entry.FieldCount(2) + entry.FieldCount(3)) {
                complete = false;
                references.Declarations.Record(owner, EntryPath(entryIndex), entry.TotalFieldCount,
                    IWorkSourceDeclarationIssueKind.RejectedMessageSet);
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
            ulong? level = entry.GetUnsigned(2);
            if (entry.FieldCount(2) != 1 || entry.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
                || !level.HasValue || level.Value > int.MaxValue) {
                complete = false;
                references.Declarations.Record(owner, EntryPath(entryIndex) + "/2", entry.FieldCount(2),
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                level = null;
            }
            // Keep an invalid value as a boundary: carrying a preceding level through
            // this declaration would silently assign its later paragraphs to that level.
            result.Add(new AttributeBoundary((int)rawIndex.Value, level, level.HasValue, entryIndex));
            ulong? unassessedValue = entry.GetUnsigned(3);
            if (entry.FieldCount(3) != 1 || entry.HasUnexpectedWireKind(3, IWorkWireKind.Varint)
                || !unassessedValue.HasValue || unassessedValue.Value > uint.MaxValue) {
                complete = false;
                references.Declarations.Record(owner, EntryPath(entryIndex) + "/3", entry.FieldCount(3),
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            } else if (unassessedValue.Value != 0) {
                complete = false;
                references.Declarations.Record(owner, EntryPath(entryIndex) + "/3", 1,
                    IWorkSourceDeclarationIssueKind.UnsupportedField);
            }
        }
        AttributeBoundary[] ordered = result.OrderBy(boundary => boundary.Index).ToArray();
        for (int index = 1; index < ordered.Length; index++) {
            if (ordered[index - 1].Index == ordered[index].Index) {
                complete = false;
                references.Declarations.Record(owner, EntryPath(ordered[index].Position), 1,
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            }
        }
        return ordered;

        string EntryPath(int position) => tablePath + "/1["
            + position.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
    }

}
