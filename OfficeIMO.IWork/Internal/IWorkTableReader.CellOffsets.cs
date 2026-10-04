namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    /// <summary>Bounds selected cells by every declared physical record without materializing out-of-table cells.</summary>
    private static bool TryReadCellLimits(IWorkSourceDocument source, IWorkWireMessage row,
        IWorkArchiveRecord tile, int position, int columns, int bufferLength, byte[] offsets,
        bool wide, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction,
        out int availableColumns, out Dictionary<int, int> cellLimits) {
        int slotCount = offsets.Length / 2;
        availableColumns = Math.Min(columns, slotCount);
        cellLimits = new Dictionary<int, int>();
        // The tile stride bounds rows, not columns. Permit supported wide tables,
        // but reject oversized trailing envelopes before inspecting their slots.
        if (slotCount > Math.Max(columns, TileRowStride)) {
            RecordInvalidCellOffsets(row, tile, position, references, diagnostics, ref supportsEditableReconstruction);
            return false;
        }
        budget.AddTableDimensionEntries(slotCount);
        var selectedOffsets = new HashSet<int>();
        var physicalOffsets = new HashSet<int>();
        int selectedCount = 0;
        bool duplicateSelectedOffset = false;
        bool populatedTrailingOffset = false;
        for (int column = 0; column < slotCount; column++) {
            source.CancellationToken.ThrowIfCancellationRequested();
            int encoded = offsets[column * 2] | offsets[column * 2 + 1] << 8;
            if (encoded == ushort.MaxValue) continue;
            int offset = wide ? encoded * 4 : encoded;
            physicalOffsets.Add(offset);
            if (column < availableColumns) {
                selectedCount++;
                if (!selectedOffsets.Add(offset)) duplicateSelectedOffset = true;
            } else populatedTrailingOffset = true;
        }
        if (populatedTrailingOffset)
            RecordInvalidCellOffsets(row, tile, position, references, diagnostics, ref supportsEditableReconstruction);
        AssessRowStorageSelection(row, tile, position, bufferLength, selectedCount,
            references, diagnostics, ref supportsEditableReconstruction);
        if (duplicateSelectedOffset) {
            RecordInvalidTileRow(tile, position, references);
            MarkCellStorageUnsupported(tile, diagnostics, ref supportsEditableReconstruction);
            return false;
        }
        int[] orderedOffsets = physicalOffsets.ToArray();
        Array.Sort(orderedOffsets);
        for (int index = 0; index < orderedOffsets.Length; index++) {
            int offset = orderedOffsets[index];
            if (!selectedOffsets.Contains(offset)) continue;
            // An invalid later offset must not extend a preceding record beyond
            // the buffer. Its own selected cell still reports a decode failure.
            cellLimits.Add(offset, index + 1 < orderedOffsets.Length
                ? Math.Min(orderedOffsets[index + 1], bufferLength) : bufferLength);
        }
        return true;
    }

    private static void RecordInvalidCellOffsets(IWorkWireMessage row, IWorkArchiveRecord tile,
        int position, IWorkSourceReferenceIssueCollector references, List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction) {
        references.Declarations.Record(tile, TileRowPath(position) + "/7", row.FieldCount(7),
            IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
        MarkCellStorageUnsupported(tile, diagnostics, ref supportsEditableReconstruction);
    }
}
