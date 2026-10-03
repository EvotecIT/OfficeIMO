using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private const uint HeaderStorageBucketArchive = 6006;

    private static void ReadHeaderDimensions(IWorkSourceDocument source, IWorkWireMessage store,
        IWorkArchiveRecord model, int rows, int columns, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references,
        Dictionary<int, double> rowHeights, Dictionary<int, double> columnWidths,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        bool complete = true;
        IWorkWireMessage? rowHeaders = IWorkObjectIndex.TryGetMessage(store, 1, out bool malformedRows);
        if (malformedRows) references.Declarations.Record(model, "4/1", store.FieldCount(1));
        complete &= !malformedRows;
        if (rowHeaders != null) {
            budget.AddTableDimensionEntries(rowHeaders.FieldCount(2));
            IReadOnlyList<IWorkArchiveRecord> buckets = references.ReadAll(model, rowHeaders, 2,
                out int unresolved, "4/1/2", static type => type == HeaderStorageBucketArchive);
            complete &= unresolved == 0;
            bool rowIndicesComplete = unresolved == 0;
            var seenRows = new HashSet<int>();
            var seenBuckets = new Dictionary<ulong, HashSet<int>>();
            foreach (IWorkArchiveRecord bucket in buckets) {
                source.CancellationToken.ThrowIfCancellationRequested();
                if (seenBuckets.TryGetValue(bucket.Identifier, out HashSet<int>? repeatedIndices)) {
                    complete = false;
                    references.Declarations.Record(model, "4/1/2", rowHeaders.FieldCount(2),
                        IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                    // Known repeats retire only their indexes. Clear the cached set after the first
                    // repeat so further references cannot multiply removal work or restore a size.
                    foreach (int position in repeatedIndices) {
                        source.CancellationToken.ThrowIfCancellationRequested();
                        rowHeights.Remove(position);
                    }
                    repeatedIndices.Clear();
                    continue;
                }
                var selectedIndices = new HashSet<int>();
                seenBuckets.Add(bucket.Identifier, selectedIndices);
                ReadDimensionBucket(source, bucket, rows, budget, seenRows, rowHeights, references,
                    selectedIndices, ref rowIndicesComplete, ref complete);
            }
            // An unreadable index may conceal a duplicate in any selected bucket on this axis.
            if (!rowIndicesComplete) rowHeights.Clear();
        }
        if (store.HasField(2)) {
            budget.AddTableDimensionEntries(store.FieldCount(2));
            IWorkArchiveRecord? bucket = references.ReadOne(model, store, 2, "4/2", static type => type == HeaderStorageBucketArchive);
            if (bucket == null) complete = false;
            else {
                bool columnIndicesComplete = true;
                ReadDimensionBucket(source, bucket, columns, budget, new HashSet<int>(), columnWidths,
                    references, null, ref columnIndicesComplete, ref complete);
                if (!columnIndicesComplete) columnWidths.Clear();
            }
        }
        if (!complete) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_HEADER_DIMENSIONS_UNSUPPORTED",
                "An iWork table has malformed, unresolved, duplicate, or hidden row/column sizing records; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
        }
    }

    private static void ReadDimensionBucket(IWorkSourceDocument source, IWorkArchiveRecord bucket,
        int dimensionCount, IWorkProjectionBudget budget, HashSet<int> seen,
        Dictionary<int, double> sizes, IWorkSourceReferenceIssueCollector references,
        HashSet<int>? selectedIndices, ref bool indicesComplete, ref bool complete) {
        if (bucket.MessageType != HeaderStorageBucketArchive) {
            complete = indicesComplete = false;
            references.Declarations.Record(bucket, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            return;
        }
        int declaredHeaders;
        try {
            declaredHeaders = IWorkProtobuf.CountFields(bucket.Payload, 2, source.Options.MaximumProtobufFieldCount);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            complete = indicesComplete = false;
            references.Declarations.Record(bucket, "$", null);
            return;
        }
        // Charge the declared headers before materializing nested messages, including invalid entries.
        budget.AddTableDimensionEntries(declaredHeaders);
        IWorkWireMessage message = source.Index.Message(bucket);
        int entryPosition = 0;
        foreach (IWorkWireValue value in message.EnumerateValues(2)) {
            source.CancellationToken.ThrowIfCancellationRequested();
            string path = "2[" + (++entryPosition).ToString(CultureInfo.InvariantCulture) + "]";
            IWorkWireMessage? header = null;
            if (value.Kind == IWorkWireKind.Bytes && value.Bytes != null) {
                try { header = message.ParseNestedMessage(value.Bytes); }
                catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) { }
            }
            if (header == null) {
                complete = indicesComplete = false;
                references.Declarations.Record(bucket, path, 1);
                continue;
            }
            ulong? index = header.GetUnsigned(1);
            float? size = header.GetFloat(2);
            ulong? hidingState = header.GetUnsigned(3);
            if (header.FieldCount(1) != 1 || header.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || index == null) {
                complete = indicesComplete = false;
                references.Declarations.Record(bucket, path + "/1", header.FieldCount(1),
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                continue;
            }
            if (index.Value >= (ulong)dimensionCount) {
                complete = false;
                references.Declarations.Record(bucket, path + "/1", 1,
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                continue;
            }
            int position = (int)index.Value + 1;
            selectedIndices?.Add(position);
            bool duplicate = !seen.Add(position);
            if (duplicate) {
                references.Declarations.Record(bucket, path + "/1", 1,
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            }
            bool invalidSize = header.FieldCount(2) != 1
                || header.HasUnexpectedWireKind(2, IWorkWireKind.Fixed32)
                || size == null || !IsFinite(size.Value) || size.Value < 0;
            bool invalidHidingState = header.FieldCount(3) != 1
                || header.HasUnexpectedWireKind(3, IWorkWireKind.Varint) || hidingState != 0;
            if (invalidSize) references.Declarations.Record(bucket, path + "/2", header.FieldCount(2),
                IWorkSourceDeclarationIssueKind.InvalidValue);
            if (invalidHidingState) references.Declarations.Record(bucket, path + "/3", header.FieldCount(3),
                IWorkSourceDeclarationIssueKind.InvalidValue);
            if (duplicate || invalidSize || invalidHidingState) {
                complete = false;
                sizes.Remove(position);
                continue;
            }
            // Native zero means use the table default, rather than a hidden or zero-height dimension.
            if (size.HasValue && size.Value > 0) sizes.Add(position, size.Value);
        }
    }
}
