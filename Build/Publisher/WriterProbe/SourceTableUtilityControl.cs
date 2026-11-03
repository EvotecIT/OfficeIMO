using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;
using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Derives the missing-utility input from the independent decorated/bare pair control.</summary>
internal static class SourceTableUtilityControl {
    internal static int Prepare(string input, string output) {
        if (File.Exists(output)) throw new IOException("Existing comparison evidence is not overwritten.");
        byte[] source = File.ReadAllBytes(input);
        if (!OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error)) throw new InvalidDataException(error);
        PublisherTable[] tables = PublisherDocument.Load(source).Pages[0].Tables.OrderBy(table => table.Y).ToArray();
        if (tables.Length != 2 || tables[0].BorderSegments.Count == 0 || tables[1].BorderSegments.Count != 0
            || tables[1].RowCount != 2 || tables[1].ColumnCount != 2 || tables[1].Cells.Count != 4
            || tables[1].Cells.Any(cell => cell.RowSpan != 1 || cell.ColumnSpan != 1 || cell.FillColor.HasValue))
            throw new InvalidDataException("Expected the independent decorated/bare two-table control.");
        byte[] art = (byte[])compound!.Streams["Escher/EscherStm"].Clone();
        var reader = new PublisherEscherReader(new PublisherBinaryData(art, "Escher/EscherStm"), null, Context());
        PublisherTableArtwork[] owners = reader.Read().TableArtwork.Where(owner => owner.TableId == tables[1].Id).ToArray();
        if (owners.Length != 4 || owners.Any(owner => owner.Value(0x2001) != 0))
            throw new InvalidDataException("Expected four hidden second-table cell records.");
        var top = reader.Records(0, art.Length); var dgg = top.Single(item => item.Kind == 0xF000);
        var fdgg = reader.Records(dgg.Offset, dgg.End).Single(item => item.Kind == 0xF006);
        var drawing = top.Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 2);
        var fdg = reader.Records(drawing.Offset, drawing.End).Single(item => item.Kind == 0xF008);
        uint oldDrawing = BitConverter.ToUInt32(art, fdg.Offset), oldTotal = BitConverter.ToUInt32(art, fdgg.Offset + 8);
        if (oldDrawing < 5 || oldTotal < oldDrawing) throw new InvalidDataException("Invalid source counts.");
        BitConverter.GetBytes(oldDrawing - 4).CopyTo(art, fdg.Offset);
        BitConverter.GetBytes(oldTotal - 4).CopyTo(art, fdgg.Offset + 8);
        var edits = new PublisherEscherShapeEdits();
        foreach (var owner in owners) edits.Removed.Add(owner.Shape.NativeBinding!.BodyOffset);
        byte[] rewritten = PublisherEscherSourceWriter.Rewrite(art, new Dictionary<int, IReadOnlyDictionary<ushort, OfficeArtProperty>>(),
            null, Context(), 64L * 1024 * 1024, edits) ?? throw new InvalidDataException("Utility removal failed.");
        byte[] result = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = rewritten });
        using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write)) stream.Write(result);
        Console.WriteLine(JsonSerializer.Serialize(new {
            sourceSha256 = Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant(),
            outputSha256 = Convert.ToHexString(SHA256.HashData(result)).ToLowerInvariant(),
            firstTableId = tables[0].Id, secondTableId = tables[1].Id, removedCellUtilities = 4,
            retainedAllocationCursor = BitConverter.ToUInt32(art, fdg.Offset + 4),
            activeDrawingCount = oldDrawing - 4, activeTotal = oldTotal - 4,
            scope = "Derived source input; native control remains independently authored. Cursors, used slots, other records and logical streams retained."
        }));
        return 0;
    }

    private static PublisherParseContext Context() => new(new PublisherReadOptions(), CancellationToken.None);
}
