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
        complete &= !malformedRows;
        if (rowHeaders != null) {
            budget.AddTableDimensionEntries(rowHeaders.FieldCount(2));
            IReadOnlyList<IWorkArchiveRecord> buckets = references.ReadAll(model, rowHeaders, 2,
                out int unresolved, "4/1/2");
            complete &= unresolved == 0;
            var seenRows = new HashSet<int>();
            var seenBuckets = new HashSet<ulong>();
            foreach (IWorkArchiveRecord bucket in buckets) {
                source.CancellationToken.ThrowIfCancellationRequested();
                if (!seenBuckets.Add(bucket.Identifier)) {
                    complete = false;
                    continue;
                }
                ReadDimensionBucket(source, bucket, rows, budget, seenRows, rowHeights, ref complete);
            }
        }
        if (store.HasField(2)) {
            budget.AddTableDimensionEntries(store.FieldCount(2));
            IWorkArchiveRecord? bucket = references.ReadOne(model, store, 2, "4/2");
            if (bucket == null) complete = false;
            else ReadDimensionBucket(source, bucket, columns, budget, new HashSet<int>(), columnWidths,
                ref complete);
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
        Dictionary<int, double> sizes, ref bool complete) {
        if (bucket.MessageType != HeaderStorageBucketArchive) {
            complete = false;
            return;
        }
        IWorkWireMessage message;
        try {
            message = source.Index.Message(bucket);
        } catch (InvalidDataException) {
            complete = false;
            return;
        }
        // Charge the declared headers before materializing nested messages, including invalid entries.
        budget.AddTableDimensionEntries(message.FieldCount(2));
        IReadOnlyList<IWorkWireMessage> headers = IWorkObjectIndex.TryGetMessages(message, 2,
            out bool malformedHeaders);
        complete &= !malformedHeaders;
        foreach (IWorkWireMessage header in headers) {
            source.CancellationToken.ThrowIfCancellationRequested();
            ulong? index = header.GetUnsigned(1);
            float? size = header.GetFloat(2);
            ulong? hidingState = header.GetUnsigned(3);
            if (header.FieldCount(1) != 1 || header.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || index == null || index.Value >= (ulong)dimensionCount) {
                complete = false;
                continue;
            }
            int position = (int)index.Value + 1;
            bool duplicate = !seen.Add(position);
            if (duplicate || header.FieldCount(2) != 1
                || header.HasUnexpectedWireKind(2, IWorkWireKind.Fixed32)
                || size == null || !IsFinite(size.Value) || size.Value < 0
                || header.FieldCount(3) != 1
                || header.HasUnexpectedWireKind(3, IWorkWireKind.Varint) || hidingState != 0) {
                complete = false;
                sizes.Remove(position);
                continue;
            }
            // Native zero means use the table default, rather than a hidden or zero-height dimension.
            if (size.Value > 0) sizes.Add(position, size.Value);
        }
    }
}
