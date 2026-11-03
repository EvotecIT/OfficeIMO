using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;
using System.Text.Json;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Derives two missing merged-anchor utilities from the independent borders-only native control.</summary>
internal static class SourceMergedTableUtilityControl {
    internal static int Prepare(string input, string output) {
        if (File.Exists(output)) throw new IOException("Existing comparison evidence is not overwritten.");
        byte[] source;
        using (var stream = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read)) {
            if (stream.Length > 64L * 1024 * 1024) throw new InvalidDataException("Probe input exceeds 64 MiB.");
            source = new byte[checked((int)stream.Length)]; stream.ReadExactly(source);
        }
        if (!OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error)) throw new InvalidDataException(error);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable[] tables = document.Pages.Concat(document.MasterPages).SelectMany(page => page.Tables).ToArray();
        (int Row, int Column, int RowSpan, int ColumnSpan)[] expected = {
            (0, 0, 1, 2), (0, 2, 1, 1), (0, 3, 1, 1), (1, 0, 2, 1), (1, 1, 1, 1), (2, 1, 1, 1), (1, 2, 2, 2)
        };
        if (document.Pages.Count != 1 || document.Pages[0].Tables.Count != 1 || tables.Length != 1)
            throw new InvalidDataException("Expected one physical page with one merged table and no other tables.");
        PublisherTable table = tables[0];
        if (table.RowCount != 3 || table.ColumnCount != 4 || !table.HasPlacement || !table.HasTextMapping
            || table.BorderSegments.Count != 19 || table.Cells.Any(cell => cell.FillColor.HasValue)
            || !table.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan)).SequenceEqual(expected))
            throw new InvalidDataException("Expected the seven native merged definitions, no cell fills and nineteen visible borders.");
        byte[] art = (byte[])compound!.Streams["Escher/EscherStm"].Clone();
        var reader = Reader(art); var before = Parts(art, reader);
        PublisherTableArtwork[] artwork = reader.Read().TableArtwork.ToArray();
        PublisherTableArtwork[] cells = artwork.Where(owner => owner.Value(0x2001) == 0).ToArray();
        var data = new PublisherBinaryData(art, "Escher/EscherStm");
        if (artwork.Length != 26 || artwork.Any(owner => owner.TableId != table.Id || owner.Value(0x2001) > 2)
            || cells.Length != 7 || !cells.Select(owner => owner.Value(0x2003)).OrderBy(index => index).SequenceEqual(Enumerable.Range(0, 7).Select(index => (uint)index))
            || cells.Any(owner => owner.Shape.NativeBinding is not { } binding || !binding.Matches(data)
                || !before.Shapes.Any(shape => shape.Offset == binding.BodyOffset) || owner.Shape.Style.FillEnabled != false
                || owner.Shape.Style.LineEnabled != false || owner.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2003)))
            || cells.Select(owner => owner.Shape.NativeBinding!.BodyOffset).Distinct().Count() != 7)
            throw new InvalidDataException("Expected seven distinct hidden cell owners bound directly to the sole table's utility drawing.");
        PublisherTableArtwork[] selected = cells.Where(owner => owner.Value(0x2003) is 0 or 6).ToArray();
        var removedBodies = selected.Select(owner => owner.Shape.NativeBinding!.BodyOffset).ToHashSet();
        uint sourceDrawingCount = data.U32(before.Descriptor.Offset), sourceTotal = data.U32(before.Allocation.Offset + 8);
        byte[] history = art.AsSpan(before.Allocation.Offset, before.Allocation.Length).ToArray(); Array.Clear(history, 8, 4);
        var removal = SourceTableUtilityControl.RemoveCellUtilities(compound, art, reader, selected);
        var outputReader = Reader(removal.Artwork); var after = Parts(removal.Artwork, outputReader);
        byte[] retainedHistory = removal.Artwork.AsSpan(after.Allocation.Offset, after.Allocation.Length).ToArray(); Array.Clear(retainedHistory, 8, 4);
        if (removal.Bytes.Length > 64L * 1024 * 1024 || !history.AsSpan().SequenceEqual(retainedHistory)
            || !Hash(art, before.Descriptor).Equals(Hash(removal.Artwork, after.Descriptor), StringComparison.Ordinal)
            || !before.Top.Where(item => item != before.Drawing).Select(item => Hash(art, item)).SequenceEqual(after.Top.Where(item => item != after.Drawing).Select(item => Hash(removal.Artwork, item)))
            || !before.Shapes.Where(item => !removedBodies.Contains(item.Offset)).Select(item => Hash(art, item)).SequenceEqual(after.Shapes.Select(item => Hash(removal.Artwork, item)))
            || BitConverter.ToUInt32(removal.Artwork, after.Descriptor.Offset) != sourceDrawingCount - 2
            || BitConverter.ToUInt32(removal.Artwork, after.Allocation.Offset + 8) != sourceTotal - 2)
            throw new InvalidDataException("Merged-anchor derivation changed retained containers, allocation history, cursor or active counts.");
        if (!OfficeCompoundFileReader.TryRead(removal.Bytes, out OfficeCompoundFile? result, out error)) throw new InvalidDataException(error);
        if (compound.Streams.Count != result!.Streams.Count || compound.Streams.Any(item => !result.Streams.TryGetValue(item.Key, out byte[]? bytes)
            || item.Key != "Escher/EscherStm" && !item.Value.AsSpan().SequenceEqual(bytes)))
            throw new InvalidDataException("Merged-anchor derivation changed another logical stream.");
        using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write)) stream.Write(removal.Bytes);
        Console.WriteLine(JsonSerializer.Serialize(new {
            schema = "officeimo.publisher.writer-probe.table-merged-source.v1", sourceSha256 = NativeInventory.Hash(source), outputSha256 = NativeInventory.Hash(removal.Bytes),
            tableId = table.Id, rows = 3, columns = 4, logicalDefinitions = table.Cells.Select((cell, index) => new {
                definitionIndex = index, cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan, removedUtility = index is 0 or 6
            }).ToArray(), removedCellIndices = new[] { 0, 6 }, sourceCellUtilities = 7, retainedCellUtilities = 5, retainedEdgeUtilities = 19,
            sourceActiveDrawingCount = sourceDrawingCount, activeDrawingCount = removal.ActiveDrawingCount, sourceActiveTotal = sourceTotal, activeTotal = removal.ActiveTotal,
            retainedAllocationCursor = removal.RetainedAllocationCursor, retainedAllocationMaximum = BitConverter.ToUInt32(history, 0),
            retainedAllocationSlots = (history.Length - 16) / 8, allocationHistorySha256 = NativeInventory.Hash(history), nativeAcceptance = "not-run",
            scope = "Derived source input only; native control remains independently authored. Only hidden cell utilities 0 and 6, enclosing lengths and two active counts change. All seven logical definitions, stories, other cell and edge containers, FDG cursor, IDCL used slots/maximum and other logical streams are retained."
        }));
        return 0;
    }

    private static (PublisherEscherRecord[] Top, PublisherEscherRecord Drawing, PublisherEscherRecord Descriptor, PublisherEscherRecord Allocation, PublisherEscherRecord[] Shapes) Parts(byte[] bytes, PublisherEscherReader reader) {
        var top = reader.Records(0, bytes.Length).ToArray();
        var drawing = top.Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(bytes, item.Offset - 12) == 2);
        var children = reader.Records(drawing.Offset, drawing.End); var group = children.Single(item => item.Kind == 0xF003);
        var shapes = reader.Records(group.Offset, group.End).ToArray(); var globals = top.Single(item => item.Kind == 0xF000);
        if (children.Count != 2 || shapes.Any(item => item.Kind != 0xF004)) throw new InvalidDataException("Expected direct utility containers and one drawing descriptor.");
        return (top, drawing, children.Single(item => item.Kind == 0xF008), reader.Records(globals.Offset, globals.End).Single(item => item.Kind == 0xF006), shapes);
    }
    private static string Hash(byte[] bytes, PublisherEscherRecord record) => NativeInventory.Hash(bytes.AsSpan(record.Offset - 8, record.Length + 8).ToArray());
    private static PublisherEscherReader Reader(byte[] bytes) => new(new PublisherBinaryData(bytes, "Escher/EscherStm"), null,
        new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
}
