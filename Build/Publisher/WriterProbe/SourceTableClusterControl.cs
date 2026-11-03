using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;
using System.Text.Json;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Derives bounded hole and synthetic compact inputs from independently authored dense native controls.</summary>
internal static class SourceTableClusterControl {
    private const int Size = 18, CellCount = Size * Size, EdgeCount = 2 * Size * (Size + 1), RemovedCount = 20;

    internal static int Prepare(string input, string output) {
        if (File.Exists(output)) throw new IOException("Existing comparison evidence is not overwritten.");
        if (new FileInfo(input).Length > 64L * 1024 * 1024) throw new InvalidDataException("Probe input exceeds 64 MiB.");
        byte[] source = File.ReadAllBytes(input);
        if (!OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error)) throw new InvalidDataException(error);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = RequireTable(document);
        byte[] art = (byte[])compound!.Streams["Escher/EscherStm"].Clone();
        var reader = Reader(art);
        PublisherEscherData artwork = reader.Read();
        AllocationSnapshot before = ReadAllocation(art, reader);
        PublisherTableArtwork[] owners = RequireOwners(table, artwork, art, before.UtilityBodies);
        int[] selected = table.Cells.Select((cell, index) => (cell, index))
            .Where(item => (item.cell.RowIndex + item.cell.ColumnIndex) % 3 == 2).Take(RemovedCount).Select(item => item.index).ToArray();
        if (selected.Length != RemovedCount || selected[0] != 2)
            throw new InvalidDataException("The dense control must begin its invisible-cell definitions at index two.");
        var removal = SourceTableUtilityControl.RemoveCellUtilities(compound, art, reader,
            selected.Select(index => owners.Single(owner => owner.Value(0x2003) == index)).ToArray());
        var outputReader = Reader(removal.Artwork);
        AllocationSnapshot after = ReadAllocation(removal.Artwork, outputReader);
        PublisherTableArtwork[] retained = outputReader.Read().TableArtwork.ToArray();
        if (after.DrawingOwner != before.DrawingOwner || after.Cursor != before.Cursor
            || !before.History.AsSpan().SequenceEqual(after.History)
            || after.ActiveDrawingCount != before.ActiveDrawingCount - RemovedCount || after.ActiveTotal != before.ActiveTotal - RemovedCount
            || retained.Length != CellCount + EdgeCount - RemovedCount
            || retained.Any(owner => owner.TableId != table.Id || owner.Value(0x2001) == 0 && selected.Contains((int)owner.Value(0x2003))))
            throw new InvalidDataException("Utility removal changed allocation history or failed to remove the selected owners exactly.");
        RequireOtherStreams(compound, removal.Bytes);
        using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write)) stream.Write(removal.Bytes);
        Console.WriteLine(JsonSerializer.Serialize(new {
            schema = "officeimo.publisher.writer-probe.table-cluster-source.v1",
            profile = "independent-native-18x18-twenty-invisible-cell-holes",
            sourceSha256 = NativeInventory.Hash(source), outputSha256 = NativeInventory.Hash(removal.Bytes),
            tableId = table.Id, rows = Size, columns = Size,
            sourceCellUtilities = CellCount, sourceEdgeUtilities = EdgeCount,
            removedCellUtilities = RemovedCount, removedCellIndices = selected,
            retainedCellUtilities = CellCount - RemovedCount, retainedEdgeUtilities = EdgeCount,
            drawingOwner = before.DrawingOwner, retainedAllocationCursor = after.Cursor,
            sourceActiveDrawingCount = before.ActiveDrawingCount, activeDrawingCount = after.ActiveDrawingCount,
            sourceActiveTotal = before.ActiveTotal, activeTotal = after.ActiveTotal,
            retainedAllocationMaximum = BitConverter.ToUInt32(after.History, 0),
            retainedAllocationSlots = (after.History.Length - 16) / 8,
            allocationHistorySha256 = NativeInventory.Hash(after.History),
            nativeAcceptance = "not-run",
            scope = "Derived source input; native control remains independently authored. Only selected invisible cell containers, enclosing lengths and active counts change; allocation history and other logical streams are retained."
        }));
        return 0;
    }

    /// <summary>Creates a synthetic initializer input by deliberately discarding utility scopes and their allocation history.</summary>
    internal static int PrepareCompact(string input, string output) {
        const int size = 20;
        if (File.Exists(output)) throw new IOException("Existing comparison evidence is not overwritten.");
        if (new FileInfo(input).Length > 64L * 1024 * 1024) throw new InvalidDataException("Probe input exceeds 64 MiB.");
        byte[] source = File.ReadAllBytes(input);
        if (!OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error)) throw new InvalidDataException(error);
        PublisherTable table = RequireTable(PublisherDocument.Load(source), size);
        byte[] art = compound!.Streams["Escher/EscherStm"];
        var reader = Reader(art);
        AllocationSnapshot before = ReadAllocation(art, reader);
        RequireOwners(table, reader.Read(), art, before.UtilityBodies, size);
        var top = reader.Records(0, art.Length);
        var dgg = top.Single(item => item.Kind == 0xF000);
        var globals = reader.Records(dgg.Offset, dgg.End);
        var allocation = globals.Single(item => item.Kind == 0xF006);
        var ordinary = top.Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 0);
        var fdg = reader.Records(ordinary.Offset, ordinary.End).Single(item => item.Kind == 0xF008);
        uint cursor = BitConverter.ToUInt32(art, fdg.Offset + 4), used = BitConverter.ToUInt32(art, allocation.Offset + 20);
        if (top.Count != 4 || top[0] != dgg || top[top.Count - 1] != ordinary || fdg.Initial != 0x10
            || BitConverter.ToUInt32(art, fdg.Offset) != 3 || cursor >> 10 != 1 || used != (cursor & 1023) + 1)
            throw new InvalidDataException("Synthetic compact derivation requires the native three-shape ordinary drawing in cluster one.");
        byte[] compactAllocation = new byte[32];
        BitConverter.GetBytes((ushort)0xF006).CopyTo(compactAllocation, 2);
        uint[] values = { 24, checked(cursor + 1), 2, 3, 1, 1, used };
        for (int index = 0; index < values.Length; index++) BitConverter.GetBytes(values[index]).CopyTo(compactAllocation, 4 + index * 4);
        var edits = new PublisherEscherShapeEdits {
            AllocationRecord = new(allocation.Offset - 8, allocation.Length + 8, compactAllocation)
        };
        byte[] compact = PublisherEscherSourceWriter.Rewrite(art, new Dictionary<int, IReadOnlyDictionary<ushort, OfficeArtProperty>>(),
            null, new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None), 64L * 1024 * 1024, edits)
            ?? throw new InvalidDataException("Synthetic compact allocation replacement failed.");
        var compactReader = Reader(compact);
        var compactTop = compactReader.Records(0, compact.Length);
        var removedDrawings = compactTop.Where(item => item.Kind == 0xF002 && item.Offset >= 12
            && BitConverter.ToUInt32(compact, item.Offset - 12) is 1 or 2).ToArray();
        if (removedDrawings.Length != 2 || removedDrawings.Select(item => BitConverter.ToUInt32(compact, item.Offset - 12)).Distinct().Count() != 2)
            throw new InvalidDataException("Expected exactly one utility and one auxiliary drawing to discard.");
        compact = PublisherNativeStreamEdit.Apply(compact, removedDrawings.Select(item =>
            new PublisherNativeStreamEdit(item.Offset - 12, item.Length + 12, Array.Empty<byte>())).ToArray(), 64L * 1024 * 1024, CancellationToken.None);
        compactReader = Reader(compact);
        compactTop = compactReader.Records(0, compact.Length);
        var keptOrdinary = compactTop.Last();
        var keptGlobals = compactReader.Records(compactTop.First().Offset, compactTop.First().End);
        if (compactTop.Count != 2 || keptOrdinary.Kind != 0xF002 || BitConverter.ToUInt32(compact, keptOrdinary.Offset - 12) != 0
            || !art.AsSpan(ordinary.Offset - 8, ordinary.Length + 8).SequenceEqual(compact.AsSpan(keptOrdinary.Offset - 8, keptOrdinary.Length + 8))
            || !globals.Where(item => item.Kind != 0xF006).Select(item => NativeInventory.Hash(art.AsSpan(item.Offset - 8, item.Length + 8).ToArray()))
                .SequenceEqual(keptGlobals.Where(item => item.Kind != 0xF006).Select(item => NativeInventory.Hash(compact.AsSpan(item.Offset - 8, item.Length + 8).ToArray()))))
            throw new InvalidDataException("Synthetic compact derivation changed retained ordinary or global records.");
        byte[] result = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = compact });
        RequireOtherStreams(compound, result);
        RequireCompactTable(table, result);
        using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write)) stream.Write(result);
        Console.WriteLine(JsonSerializer.Serialize(new {
            schema = "officeimo.publisher.writer-probe.table-cluster-compact-source.v1", profile = "synthetic-compact-20x20-initializer-input",
            sourceSha256 = NativeInventory.Hash(source), outputSha256 = NativeInventory.Hash(result), tableId = table.Id, rows = size, columns = size,
            removedCellUtilities = size * size, removedEdgeUtilities = 2 * size * (size + 1), removedDrawingInstances = new[] { 2, 1 },
            retainedDrawingOwner = 1, retainedAllocationCursor = cursor, retainedOrdinaryUsed = used,
            activeDrawingCount = 3, activeTotal = 3, allocationMaximum = cursor + 1, allocationClusterCount = 2, savedDrawingCount = 1,
            sourceAllocationHistorySha256 = NativeInventory.Hash(before.History), nativeAcceptance = "not-run", noOpExact = true,
            scope = "Synthetic compact source/control only. All table paint and utility/auxiliary allocation history are deliberately discarded to qualify initializer creation; this is not a preserving user operation. Ordinary records, other global records and other logical streams are retained."
        }));
        return 0;
    }

    private static void RequireCompactTable(PublisherTable source, byte[] result) {
        PublisherDocument loaded = PublisherDocument.Load(result);
        PublisherTable[] tables = loaded.Pages.Concat(loaded.MasterPages).SelectMany(page => page.Tables).ToArray();
        if (tables.Length != 1) throw new InvalidDataException("Synthetic compact output lost its sole table.");
        PublisherTable table = tables[0];
        if (table.Id != source.Id || table.RowCount != source.RowCount || table.ColumnCount != source.ColumnCount || !table.HasTextMapping
            || !table.HasPlacement || table.X != source.X || table.Y != source.Y || table.Width != source.Width || table.Height != source.Height
            || !table.PageTransform.Equals(source.PageTransform) || !table.RowHeights.SequenceEqual(source.RowHeights)
            || !table.ColumnWidths.SequenceEqual(source.ColumnWidths) || table.BorderSegments.Count != 0 || table.Cells.Any(cell => cell.FillColor.HasValue)
            || !table.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, cell.X, cell.Y, cell.Width, cell.Height, cell.Text))
                .SequenceEqual(source.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, cell.X, cell.Y, cell.Width, cell.Height, cell.Text))))
            throw new InvalidDataException("Synthetic compact output changed cell text or geometry, or retained table paint.");
        PublisherNativePreparation noOp = loaded.PrepareNativeBytes(64L * 1024 * 1024);
        if (!noOp.IsPrepared || !noOp.CopyBytes().AsSpan().SequenceEqual(result))
            throw new InvalidDataException("Synthetic compact output is not an exact native no-op source.");
    }

    private static PublisherTable RequireTable(PublisherDocument document, int size = Size) {
        int cellCount = size * size, edgeCount = 2 * size * (size + 1);
        PublisherTable[] tables = document.Pages.Concat(document.MasterPages).SelectMany(page => page.Tables).ToArray();
        if (document.Pages.Count != 1 || tables.Length != 1 || document.Pages[0].Tables.Count != 1)
            throw new InvalidDataException("Expected a single-page native control with one dense table.");
        PublisherTable table = tables[0];
        if (table.RowCount != size || table.ColumnCount != size || table.Cells.Count != cellCount
            || table.BorderSegments.Count != edgeCount || !table.HasPlacement || !table.HasTextMapping
            || !double.IsFinite(table.X) || !double.IsFinite(table.Y) || !Positive(table.Width) || !Positive(table.Height)
            || table.RowHeights.Any(value => !Positive(value)) || table.ColumnWidths.Any(value => !Positive(value)))
            throw new InvalidDataException($"Expected a placed {size}-by-{size} table with {cellCount} cells and {edgeCount} visible unit edges.");
        for (int index = 0; index < cellCount; index++) {
            PublisherTableCell cell = table.Cells[index];
            int row = index / size, column = index % size;
            if (cell.RowIndex != row || cell.ColumnIndex != column || cell.RowSpan != 1 || cell.ColumnSpan != 1
                || !Close(cell.X, table.ColumnWidths.Take(column).Sum()) || !Close(cell.Y, table.RowHeights.Take(row).Sum())
                || !Close(cell.Width, table.ColumnWidths[column]) || !Close(cell.Height, table.RowHeights[row])
                || cell.FillColor.HasValue != ((row + column) % 3 != 2))
                throw new InvalidDataException("Dense-table cell geometry, definition order or three-way fill pattern differs from the control.");
        }
        return table;
    }

    private static PublisherTableArtwork[] RequireOwners(PublisherTable table, PublisherEscherData artwork, byte[] bytes, HashSet<int> bodies, int size = Size) {
        int cellCount = size * size, edgeCount = 2 * size * (size + 1);
        PublisherTableArtwork[] owners = artwork.TableArtwork.ToArray();
        if (owners.Length != cellCount + edgeCount || owners.Any(owner => owner.TableId != table.Id)
            || bodies.Count != owners.Length + 3)
            throw new InvalidDataException($"Expected all {cellCount} cell owners, {edgeCount} edge owners, root and two utility defaults.");
        var data = new PublisherBinaryData(bytes, "Escher/EscherStm");
        var cells = new HashSet<uint>();
        var edges = new HashSet<(int Row, int Column, int LastRow, int LastColumn)>();
        var ownerBodies = new HashSet<int>();
        foreach (PublisherTableArtwork owner in owners) {
            var shape = owner.Shape; var style = shape.Style;
            if (shape.NativeBinding is not { } binding || !binding.Matches(data) || !bodies.Contains(binding.BodyOffset)
                || !ownerBodies.Add(binding.BodyOffset) || shape.Type != 1 || shape.Flags != 0xA00 || shape.Hidden
                || shape.Transform.RotationDegrees.GetValueOrDefault() != 0 || shape.Transform.FlipHorizontal || shape.Transform.FlipVertical
                || style.ShadowEnabled == true || style.FillType.GetValueOrDefault() != 0 || style.FillOpacity.GetValueOrDefault(1) != 1
                || style.LineEnabled != false)
                throw new InvalidDataException("Dense-table utilities require distinct direct native bindings with simple fill-only paint.");
            uint kind = owner.Value(0x2001);
            if (kind == 0) {
                uint index = owner.Value(0x2003);
                if (index >= cellCount || !cells.Add(index) || owner.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2003)))
                    throw new InvalidDataException("Dense-table cell ownership is incomplete or ambiguous.");
                bool invisible = (table.Cells[(int)index].RowIndex + table.Cells[(int)index].ColumnIndex) % 3 == 2;
                if (invisible ? style.FillEnabled != false : style.FillEnabled == false)
                    throw new InvalidDataException("Native fill visibility differs from the independent three-way cell pattern.");
            } else if (kind is 1 or 2) {
                uint row = owner.Value(0x2004), column = owner.Value(0x2005), lastRow = owner.Value(0x2006), lastColumn = owner.Value(0x2007);
                if (row > size || column > size || lastRow > size || lastColumn > size
                    || (kind == 1 ? row != lastRow || lastColumn != column + 1 : column != lastColumn || lastRow != row + 1)
                    || !edges.Add(((int)row, (int)column, (int)lastRow, (int)lastColumn))
                    || owner.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2004 or 0x2005 or 0x2006 or 0x2007))
                    || style.FillEnabled == false || style.LineWidthEmus.GetValueOrDefault() <= 0)
                    throw new InvalidDataException("Dense-table edges require unique visible unit grid boundaries.");
            } else throw new InvalidDataException("Unexpected dense-table utility kind.");
        }
        if (cells.Count != cellCount || edges.Count != edgeCount
            || !edges.SetEquals(table.BorderSegments.Select(edge => (edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn))))
            throw new InvalidDataException("Native utility ownership does not cover the complete dense table and visible geometry.");
        return owners.Where(owner => owner.Value(0x2001) == 0).ToArray();
    }

    /// <summary>Checks live counts and SPIDs against their actual cluster owners while retaining retired allocation slots.</summary>
    private static AllocationSnapshot ReadAllocation(byte[] bytes, PublisherEscherReader reader) {
        var data = new PublisherBinaryData(bytes, "Escher/EscherStm");
        var top = reader.Records(0, bytes.Length);
        var dgg = top.Single(item => item.Kind == 0xF000);
        var allocation = reader.Records(dgg.Offset, dgg.End).Single(item => item.Kind == 0xF006);
        uint slots = data.U32(allocation.Offset + 4);
        if (allocation.Initial != 0 || slots < 2 || 16L + (slots - 1L) * 8 != allocation.Length)
            throw new InvalidDataException("Invalid dense-control allocation table.");
        for (int offset = allocation.Offset + 16; offset < allocation.End; offset += 8)
            if (data.U32(offset + 4) > 1024) throw new InvalidDataException("Invalid native allocation-slot usage.");
        var spids = new HashSet<uint>();
        var drawingIds = new HashSet<uint>();
        var utilityBodies = new HashSet<int>();
        uint total = 0, utilityOwner = 0, cursor = 0, utilityCount = 0;
        var drawings = top.Where(item => item.Kind == 0xF002).ToArray();
        foreach (var drawing in drawings) {
            var children = reader.Records(drawing.Offset, drawing.End);
            var fdg = children.Single(item => item.Kind == 0xF008);
            uint drawingId = (uint)(fdg.Initial >> 4);
            if ((fdg.Initial & 15) != 0 || fdg.Length != 8 || drawingId == 0 || !drawingIds.Add(drawingId))
                throw new InvalidDataException("Invalid native drawing descriptor.");
            uint drawingCursor = data.U32(fdg.Offset + 4);
            ValidateSlot(drawingCursor, drawingId, isCursor: true);
            uint count = CountShapes(drawing.Offset, drawing.End, drawingId, 0);
            if (count != data.U32(fdg.Offset)) throw new InvalidDataException("Native drawing active count disagrees with its shape records.");
            total = checked(total + count);
            if (drawing.Offset < 12 || data.U32(drawing.Offset - 12) != 2) continue;
            if (utilityOwner != 0 || children.Count != 2 || children.Count(item => item.Kind == 0xF003) != 1)
                throw new InvalidDataException("Expected one direct table-utility drawing.");
            var group = children.Single(item => item.Kind == 0xF003);
            var shapes = reader.Records(group.Offset, group.End);
            if (shapes.Any(item => item.Kind != 0xF004) || shapes.Count != count)
                throw new InvalidDataException("Expected direct utility shape containers.");
            utilityBodies.UnionWith(shapes.Select(shape => shape.Offset));
            utilityOwner = drawingId; cursor = drawingCursor; utilityCount = count;
        }
        if (utilityOwner == 0 || spids.Count == 0 || data.U32(allocation.Offset) <= spids.Max()
            || total != data.U32(allocation.Offset + 8) || drawings.Length != data.U32(allocation.Offset + 12))
            throw new InvalidDataException("Native global counts, allocation maximum or utility ownership are inconsistent.");
        byte[] history = bytes.AsSpan(allocation.Offset, allocation.Length).ToArray();
        Array.Clear(history, 8, 4); // The active total is the sole edited FDGG field.
        return new(utilityOwner, cursor, utilityCount, total, history, utilityBodies);

        void ValidateSlot(uint spid, uint owner, bool isCursor) {
            uint cluster = spid >> 10;
            if (cluster == 0 || cluster >= slots) throw new InvalidDataException("Live SPID lies outside the native allocation table.");
            int offset = checked(allocation.Offset + 16 + ((int)cluster - 1) * 8);
            uint used = data.U32(offset + 4);
            if (data.U32(offset) != owner || (spid & 1023) >= used || isCursor && (spid & 1023) + 1 != used)
                throw new InvalidDataException("Live SPID or cursor disagrees with its actual allocation-cluster owner.");
        }

        uint CountShapes(int start, int end, uint owner, int depth) {
            if (depth > 16) throw new InvalidDataException("Dense-control shape nesting exceeds its bounded profile.");
            uint count = 0;
            foreach (var record in reader.Records(start, end)) {
                if (record.Kind == 0xF004) {
                    var fsp = reader.Records(record.Offset, record.End).Single(item => item.Kind == 0xF00A);
                    if (fsp.Length != 8 || !spids.Add(data.U32(fsp.Offset))) throw new InvalidDataException("Native shape identifiers are missing or repeated.");
                    ValidateSlot(data.U32(fsp.Offset), owner, isCursor: false);
                    if ((data.U32(fsp.Offset + 4) & 0x400) == 0) count++;
                } else if ((record.Initial & 15) == 15 && record.Kind != 0xF001)
                    count = checked(count + CountShapes(record.Offset, record.End, owner, depth + 1));
            }
            return count;
        }
    }

    private static void RequireOtherStreams(OfficeCompoundFile source, byte[] result) {
        if (!OfficeCompoundFileReader.TryRead(result, out OfficeCompoundFile? output, out string? error)) throw new InvalidDataException(error);
        if (source.Streams.Count != output!.Streams.Count || source.Streams.Any(item => !output.Streams.TryGetValue(item.Key, out byte[]? bytes)
            || item.Key != "Escher/EscherStm" && !item.Value.AsSpan().SequenceEqual(bytes)))
            throw new InvalidDataException("Utility removal changed an unrelated logical stream.");
    }

    private static bool Positive(double value) => double.IsFinite(value) && value > 0;
    private static bool Close(double left, double right) => double.IsFinite(left) && Math.Abs(left - right) <= 1D / 12700;
    private static PublisherEscherReader Reader(byte[] bytes) => new(new PublisherBinaryData(bytes, "Escher/EscherStm"), null,
        new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
    private sealed record AllocationSnapshot(uint DrawingOwner, uint Cursor, uint ActiveDrawingCount, uint ActiveTotal,
        byte[] History, HashSet<int> UtilityBodies);
}
