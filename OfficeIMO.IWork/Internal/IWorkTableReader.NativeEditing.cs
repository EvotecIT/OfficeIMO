namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    /// <summary>Locates an existing active BNC scalar using the same tile and offset bounds as projection.</summary>
    internal static (IWorkArchiveRecord Record, int RowOccurrence, int ValueOffset, byte[] Original)
        LocateNativeScalar(IWorkSourceDocument source, IWorkTable table, IReadOnlyList<IWorkTable> tables,
            int row, int column, IWorkCellKind kind) {
        int tileIndex = (row - 1) / TileRowStride;
        var entries = NativeTileEntries(source, table).Where(entry => entry.GetUnsigned(1) == (ulong)tileIndex).Take(2).ToArray();
        IWorkArchiveRecord? tile = entries.Length == 1 ? source.Index.Dereference(entries[0], 2) : null;
        if (tile == null || tile.MessageType != TableTileArchive || tile.PayloadIndex != 0)
            throw new NotSupportedException("The selected cell must belong to one primary modern tile record.");
        int owners = 0;
        foreach (IWorkTable owner in tables) {
            source.CancellationToken.ThrowIfCancellationRequested();
            owners += NativeTileEntries(source, owner).Count(entry => source.Index.Dereference(entry, 2)?.Identifier == tile.Identifier);
        }
        if (owners != 1) throw new NotSupportedException("Shared table tiles cannot be edited independently.");

        var references = new IWorkSourceReferenceIssueCollector(source);
        var diagnostics = new List<IWorkDiagnostic>();
        var budget = new IWorkProjectionBudget(source.Options);
        bool complete = true;
        int maximumRows = Math.Min(TileRowStride, table.RowCount - tileIndex * TileRowStride);
        var message = ReadTileMessage(source, tile, maximumRows, references, diagnostics, ref complete);
        if (message == null || message.GetUnsigned(6) != 5 || message.GetUnsigned(7) != 1)
            throw new NotSupportedException("Native scalar edits require explicitly active version-five BNC storage.");
        var rows = ReadTileRows(source, tile, message, references, diagnostics, ref complete);
        var matches = rows.Where(item => item.Message.GetUnsigned(1) == (ulong)((row - 1) % TileRowStride)).Take(2).ToArray();
        if (!complete || matches.Length != 1) throw new NotSupportedException("The selected cell row must resolve unambiguously.");
        var selected = matches[0]; var rowInfo = selected.Message;
        byte[]? buffer = rowInfo.GetBytes(6), offsets = rowInfo.GetBytes(7);
        if (rowInfo.FieldCount(5) != 1 || rowInfo.HasUnexpectedWireKind(5, IWorkWireKind.Varint)
            || rowInfo.GetUnsigned(5) != 5 || buffer == null || offsets == null || offsets.Length % 2 != 0)
            throw new NotSupportedException("The selected row must declare modern scalar storage.");
        bool wide = rowInfo.GetUnsigned(8) == 1;
        if (!TryReadCellLimits(source, rowInfo, tile, selected.Position, table.ColumnCount, buffer.Length,
                offsets, wide, budget, references, diagnostics, ref complete, out int columns, out var limits)
            || !complete || column > columns)
            throw new NotSupportedException("The selected row has unsupported cell offsets.");
        int encoded = offsets[(column - 1) * 2] | offsets[(column - 1) * 2 + 1] << 8;
        int start = wide ? encoded * 4 : encoded;
        if (encoded == ushort.MaxValue || !limits.TryGetValue(start, out int end) || start > end - 12)
            throw new NotSupportedException("The selected cell has no complete existing scalar header.");
        uint flags = IWorkProtobuf.ReadUInt32(buffer, start + 8);
        bool isDecimal = (flags & 3) == 1;
        int length = isDecimal ? 16 : 8;
        if (buffer[start] != 5 || (kind == IWorkCellKind.Number ? buffer[start + 1] is not (2 or 10) : buffer[start + 1] != 6)
            || (flags & 3) is not (1 or 2) || (kind == IWorkCellKind.Number) != isDecimal
            || (flags & (1u << 9)) != 0 || (flags & ~RecognizedCellValueMask) != 0 || length > end - start - 12)
            throw new NotSupportedException("The selected cell must contain a non-formula Decimal128 number or double-backed boolean of the requested kind.");
        RequireEmptyLegacyCell(source, table, tile, rowInfo, selected.Position, column, budget, references, diagnostics);
        var original = new byte[length]; Buffer.BlockCopy(buffer, start + 12, original, 0, length);
        return (tile, selected.Position - 1, start + 12, original);
    }

    private static IReadOnlyList<IWorkWireMessage> NativeTileEntries(IWorkSourceDocument source, IWorkTable table) {
        if (table.ModelRecord == null) throw new NotSupportedException("The table must resolve to a native model.");
        var model = source.Index.Message(table.ModelRecord);
        var store = IWorkObjectIndex.TryGetMessage(model, 4);
        var tiles = store == null ? null : IWorkObjectIndex.TryGetMessage(store, 3);
        if (tiles == null) throw new NotSupportedException("The table must resolve to modern tile storage.");
        var entries = IWorkObjectIndex.TryGetMessages(tiles, 1, out bool malformed);
        if (malformed) throw new NotSupportedException("The table tile declarations must be complete.");
        return entries;
    }

    /// <summary>Apple retains empty pre-BNC placeholders; a second populated representation is not edited.</summary>
    private static void RequireEmptyLegacyCell(IWorkSourceDocument source, IWorkTable table, IWorkArchiveRecord tile,
        IWorkWireMessage row, int position, int column, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references, List<IWorkDiagnostic> diagnostics) {
        byte[]? buffer = row.GetBytes(3), offsets = row.GetBytes(4);
        if ((buffer?.Length ?? 0) == 0 && (offsets?.Length ?? 0) == 0) return;
        bool complete = true;
        if (buffer == null || offsets == null || offsets.Length % 2 != 0
            || !TryReadCellLimits(source, row, tile, position, table.ColumnCount, buffer.Length, offsets, false,
                budget, references, diagnostics, ref complete, out int columns, out var limits) || !complete || column > columns)
            throw new NotSupportedException("The legacy cell representation must be absent or an empty placeholder.");
        int start = offsets[(column - 1) * 2] | offsets[(column - 1) * 2 + 1] << 8;
        if (start == ushort.MaxValue) return;
        if (!limits.TryGetValue(start, out int end) || end - start != 12 || start < 0 || end > buffer.Length || buffer[start] != 4)
            throw new NotSupportedException("The legacy cell representation must be an empty version-four placeholder.");
        for (int index = start + 1; index < end; index++)
            if (buffer[index] != 0) throw new NotSupportedException("Populated legacy cell representations cannot be edited.");
    }
}
