using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPopulatedTableUtilityTests {
    [Theory]
    [InlineData("Styled")]
    [InlineData("Borders")]
    [InlineData("Fills")]
    [InlineData("Composed")]
    public void First_styles_in_a_populated_utility_drawing_match_independent_tables_and_preserve_peers(string variant) {
        byte[] source = MissingCells();
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable[] original = Tables(document);
        Assert.Equal(2, original.Length);
        Assert.All(original[1].Cells, cell => Assert.Null(cell.FillColor)); Assert.Empty(original[1].BorderSegments);
        Assert.Empty(Utilities(source, original[1].Id, cellsOnly: false));
        Assert.Equal(source, document.ToPubBytes());
        uint cursor = Cursor(source);
        if (variant == "Composed") PublisherTableBorderTopologyTests.Apply(document, "TableBorderTopology");
        PublisherTableUtilityCreationTests.Apply(document, variant switch {
            "Borders" => "TableBordersOnly", "Fills" => "TableFillsOnly", "Composed" => "TableDecorationText", _ => "TableDecoration"
        }, original[1]);
        var result = document.ToPubBytesResult(Allow());
        PublisherDocument loaded = PublisherDocument.Load(result.Value);
        Compare(Control(variant), loaded);
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Escher/EscherStm",
            variant == "Composed" ? new[] { "Quill/QuillSub/CONTENTS" } : Array.Empty<string>());
        Assert.Equal(Ordinary(source), Ordinary(result.Value));
        var peer = Utilities(source, original[0].Id, cellsOnly: variant == "Composed");
        var afterPeer = Utilities(result.Value, original[0].Id, cellsOnly: variant == "Composed");
        Assert.Equal(peer.Keys.OrderBy(id => id), afterPeer.Keys.OrderBy(id => id));
        foreach (var record in peer) Assert.Equal(record.Value, afterPeer[record.Key]);
        var added = Utilities(result.Value, original[1].Id, cellsOnly: false);
        Assert.Equal(variant == "Fills" ? 4 : 10, added.Count);
        Assert.All(added.Keys, id => Assert.True(id > cursor));
        Assert.True(Cursor(result.Value) >= added.Keys.Max());
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_TABLE_UTILITY_CREATION_UNQUALIFIED"
            && item.LossKind == OfficeConversionLossKind.Unassessed);
        Assert.Equal(result.Value, loaded.ToPubBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Later_edits_of_allocated_cells_preserve_peers_and_identifier_history(bool activateFills) {
        PublisherDocument document = PublisherDocument.Load(MissingCells());
        PublisherTable second = Tables(document)[1];
        PublisherTableUtilityCreationTests.Apply(document, "TableBordersOnly", second);
        byte[] bordered = document.ToPubBytes(Allow()); uint cursor = Cursor(bordered);
        document = PublisherDocument.Load(bordered); second = Tables(document)[1];
        if (activateFills) {
            PublisherTableUtilityCreationTests.Apply(document, "TableDecoration", second);
        } else foreach (var edge in second.BorderSegments) document.SetTableBorderStyle(second.Id,
            edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, null);
        byte[] edited = document.ToPubBytes(Allow());
        Compare(Control(activateFills ? "Styled" : ""), PublisherDocument.Load(edited));
        Assert.Equal(cursor, Cursor(edited));
        PublisherNativeWritingTests.AssertPreservedStorage(bordered, edited, "Escher/EscherStm");
        var beforePeer = Utilities(bordered, Tables(document)[0].Id, cellsOnly: false);
        var afterPeer = Utilities(edited, Tables(document)[0].Id, cellsOnly: false);
        foreach (var record in beforePeer) Assert.Equal(record.Value, afterPeer[record.Key]);
        if (activateFills) return;
        document = PublisherDocument.Load(edited);
        PublisherTableUtilityCreationTests.Apply(document, "TableBordersOnly", Tables(document)[1]);
        byte[] restored = document.ToPubBytes(Allow());
        Compare(Control("Borders"), PublisherDocument.Load(restored));
        Assert.True(Cursor(restored) > cursor);
        var oldIds = Utilities(bordered, second.Id, cellsOnly: false).Keys;
        var newIds = Utilities(restored, second.Id, cellsOnly: false).Keys;
        Assert.Equal(4, oldIds.Intersect(newIds).Count()); // Cell records survive; retired edges do not return.
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Missing_cell_allocation_refuses_inherited_transparency_in_populated_scopes(bool maskedByUtilityDefaults) {
        OfficeCompoundFile compound = Compound(MissingCells());
        byte[] art = (byte[])compound.Streams["Escher/EscherStm"].Clone(); var reader = Reader(art);
        var dgg = reader.Records(0, art.Length).Single(item => item.Kind == 0xF000);
        var properties = reader.Records(dgg.Offset, dgg.End).Single(item => item.Kind == 0xF00B);
        int slot = Enumerable.Range(0, properties.Initial >> 4).Select(index => properties.Offset + index * 6)
            .Single(offset => (BitConverter.ToUInt16(art, offset) & 0x3FFF) == 0x0181);
        BitConverter.GetBytes((ushort)0x0182).CopyTo(art, slot);
        BitConverter.GetBytes(32768U).CopyTo(art, slot + 2);
        if (maskedByUtilityDefaults) {
            var drawing = reader.Records(0, art.Length).Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 2);
            var group = reader.Records(drawing.Offset, drawing.End).Single(item => item.Kind == 0xF003);
            foreach (var shape in reader.Records(group.Offset, group.End).Skip(1).Take(2)) {
                var defaults = reader.Records(shape.Offset, shape.End).Single(item => item.Kind == 0xF00B);
                var entries = Enumerable.Range(0, defaults.Initial >> 4).Select(index => (
                    Id: BitConverter.ToUInt16(art, defaults.Offset + index * 6), Value: BitConverter.ToUInt32(art, defaults.Offset + index * 6 + 2)))
                    .Select(item => item.Id == 0x0081 ? (Id: (ushort)0x0182, Value: 65536U) : item).OrderBy(item => item.Id).ToArray();
                Assert.Single(entries, item => item.Id == 0x0182);
                for (int index = 0; index < entries.Length; index++) {
                    BitConverter.GetBytes(entries[index].Id).CopyTo(art, defaults.Offset + index * 6);
                    BitConverter.GetBytes(entries[index].Value).CopyTo(art, defaults.Offset + index * 6 + 2);
                }
            }
        }
        byte[] source = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(source); Assert.Equal(source, document.ToPubBytes());
        PublisherTableUtilityCreationTests.Apply(document, "TableDecoration", Tables(document)[1]);
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(Allow()));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    [Fact]
    public void New_borders_in_populated_scopes_refuse_inherited_transparency() {
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableUtilityPairFills/native-tableutilitypairfills.pub"));
        OfficeCompoundFile compound = Compound(original); byte[] art = (byte[])compound.Streams["Escher/EscherStm"].Clone();
        var reader = Reader(art); var dgg = reader.Records(0, art.Length).Single(item => item.Kind == 0xF000);
        var properties = reader.Records(dgg.Offset, dgg.End).Single(item => item.Kind == 0xF00B);
        int slot = Enumerable.Range(0, properties.Initial >> 4).Select(index => properties.Offset + index * 6)
            .Single(offset => (BitConverter.ToUInt16(art, offset) & 0x3FFF) == 0x0181);
        BitConverter.GetBytes((ushort)0x0182).CopyTo(art, slot); BitConverter.GetBytes(32768U).CopyTo(art, slot + 2);
        byte[] source = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(source); Assert.Equal(source, document.ToPubBytes());
        PublisherTable second = Tables(document)[1]; Assert.Empty(second.BorderSegments);
        PublisherTableUtilityCreationTests.Apply(document, "TableBordersOnly", second);
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(Allow()));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    [Fact]
    public void Missing_cell_allocation_refuses_exhausted_cluster_without_reusing_retired_identifiers() {
        byte[] original = MissingCells(); OfficeCompoundFile compound = Compound(original);
        byte[] art = (byte[])compound.Streams["Escher/EscherStm"].Clone(); var reader = Reader(art);
        var top = reader.Records(0, art.Length); var dgg = top.Single(item => item.Kind == 0xF000);
        var fdgg = reader.Records(dgg.Offset, dgg.End).Single(item => item.Kind == 0xF006);
        var drawing = top.Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 2);
        var fdg = reader.Records(drawing.Offset, drawing.End).Single(item => item.Kind == 0xF008);
        uint cluster = BitConverter.ToUInt32(art, fdg.Offset + 4) >> 10;
        BitConverter.GetBytes((cluster << 10) + 1023).CopyTo(art, fdg.Offset + 4);
        BitConverter.GetBytes(1024U).CopyTo(art, fdgg.Offset + 16 + ((int)cluster - 1) * 8 + 4);
        BitConverter.GetBytes((cluster + 1) << 10).CopyTo(art, fdgg.Offset);
        byte[] source = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(source); Assert.Equal(source, document.ToPubBytes());
        PublisherTableUtilityCreationTests.Apply(document, "TableDecoration", Tables(document)[1]);
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(Allow()));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    private static PublisherSaveOptions Allow() => new() { LossPolicy = OfficeConversionLossPolicy.Allow };
    private static PublisherTable[] Tables(PublisherDocument document) => document.Pages[0].Tables.OrderBy(table => table.Y).ToArray();
    private static PublisherDocument Control(string variant) => PublisherDocument.Load(PublisherNativeTests.Fixture(
        "NativeComparison/TableUtilityPair" + variant + "/native-tableutilitypair" + variant.ToLowerInvariant() + ".pub"));
    private static byte[] MissingCells() => File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableUtilityPairMissingCells/missing-second-utilities.pub"));
    private static void Compare(PublisherDocument expected, PublisherDocument actual) {
        var left = Tables(expected); var right = Tables(actual); Assert.Equal(left.Length, right.Length);
        for (int index = 0; index < left.Length; index++) {
            Assert.Equal(left[index].Cells.Select(cell => cell.Text), right[index].Cells.Select(cell => cell.Text));
            Assert.Equal(left[index].Cells.Select(cell => cell.FillColor), right[index].Cells.Select(cell => cell.FillColor));
            Assert.Equal(left[index].ColumnWidths, right[index].ColumnWidths); Assert.Equal(left[index].RowHeights, right[index].RowHeights);
            Assert.Equal((left[index].X, left[index].Y, left[index].Width, left[index].Height), (right[index].X, right[index].Y, right[index].Width, right[index].Height));
            Assert.Equal(PublisherTableBorderTopologyTests.Coverage(left[index]), PublisherTableBorderTopologyTests.Coverage(right[index]));
        }
    }
    private static Dictionary<uint, byte[]> Utilities(byte[] source, uint tableId, bool cellsOnly) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"]; var parsed = Reader(art).Read();
        return parsed.TableArtwork.Where(item => item.TableId == tableId && (!cellsOnly || item.Value(0x2001) == 0))
            .ToDictionary(item => item.Shape.NativeBinding!.Spid, item => art.Skip(item.Shape.NativeBinding!.BodyOffset - 8).Take(item.Shape.NativeBinding.BodyLength + 8).ToArray());
    }
    private static byte[] Ordinary(byte[] source) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"]; var reader = Reader(art);
        var drawing = reader.Records(0, art.Length).Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 0);
        return art.Skip(drawing.Offset - 8).Take(drawing.Length + 8).ToArray();
    }
    private static uint Cursor(byte[] source) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"]; var reader = Reader(art);
        var drawing = reader.Records(0, art.Length).Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 2);
        var fdg = reader.Records(drawing.Offset, drawing.End).Single(item => item.Kind == 0xF008);
        return BitConverter.ToUInt32(art, fdg.Offset + 4);
    }
    private static PublisherEscherReader Reader(byte[] art) => new(new PublisherBinaryData(art, "Escher/EscherStm"), null,
        new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
    private static OfficeCompoundFile Compound(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? result, out string? error), error); return result!;
    }
}
