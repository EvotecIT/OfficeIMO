using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableBorderTopologyTests {
    [Theory]
    [InlineData("TableBorderTopology")]
    [InlineData("TableBordersCleared")]
    [InlineData("TableBorderTopologyTextFill")]
    [InlineData("TableBorderDefaultFill")]
    public void Native_border_topology_matches_independent_producer_and_preserves_other_storage(string variant) {
        byte[] source = Fixture("TableDecoration");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = document.Pages[0].Tables[0];
        Apply(document, variant);
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Escher/EscherStm",
            variant == "TableBorderTopologyTextFill" ? new[] { "Quill/QuillSub/CONTENTS" } : Array.Empty<string>());
        PublisherTable actual = PublisherDocument.Load(result.Value).Pages[0].Tables[0];
        PublisherTable expected = PublisherDocument.Load(Fixture(variant)).Pages[0].Tables[0];
        Assert.Equal(Coverage(expected), Coverage(actual));
        Assert.Equal(expected.Cells.Select(cell => cell.FillColor), actual.Cells.Select(cell => cell.FillColor));
        Assert.Equal(expected.Cells.Select(cell => cell.Text), actual.Cells.Select(cell => cell.Text));
        Assert.Equal(original.ColumnWidths, actual.ColumnWidths);
        Assert.Equal(original.RowHeights, actual.RowHeights);
        Assert.Equal(result.Value, PublisherDocument.Load(result.Value).ToPubBytes());
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_TABLE_BORDER_TOPOLOGY_UNQUALIFIED" && item.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Fact]
    public void Deleting_reloading_and_readding_edges_retains_allocation_history_and_unique_shape_ids() {
        byte[] source = Fixture("TableDecoration");
        PublisherTable original = PublisherDocument.Load(source).Pages[0].Tables[0];
        PublisherDocument document = PublisherDocument.Load(source);
        Apply(document, "TableBordersCleared");
        byte[] cleared = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        var before = Allocation(source); var deleted = Allocation(cleared);
        Assert.Equal(before.Cursor, deleted.Cursor);
        Assert.Equal(before.Used, deleted.Used);
        Assert.Equal(before.Maximum, deleted.Maximum);
        Assert.Equal(before.Count - 6, deleted.Count);
        document = PublisherDocument.Load(cleared);
        foreach (var edge in original.BorderSegments) document.SetTableBorderStyle(original.Id, edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, edge.Side);
        byte[] restored = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        var added = Allocation(restored);
        Assert.Equal(before.Count, added.Count);
        Assert.Equal(before.Cursor + 6, added.Cursor);
        Assert.Equal(before.Used + 6, added.Used);
        Assert.Equal(Coverage(original), Coverage(PublisherDocument.Load(restored).Pages[0].Tables[0]));
        Assert.Equal(6, added.Ids.Except(deleted.Ids).Count());
        Assert.All(added.Ids.Except(deleted.Ids), id => Assert.True(id > before.Cursor));
        PublisherNativeWritingTests.AssertPreservedStorage(source, restored, "Escher/EscherStm");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Invalid_shape_count_full_cluster_or_cursor_below_live_ids_refuses_topology_output(int invalid) {
        byte[] source = Fixture("TableDecoration");
        OfficeCompoundFile compound = Compound(source);
        byte[] art = (byte[])compound.Streams["Escher/EscherStm"].Clone();
        var records = DrawingRecords(art);
        PublisherEscherRecord global = records.Reader.Records(records.Dgg.Offset, records.Dgg.End).Single(item => item.Kind == 0xF006);
        if (invalid != 0) {
            PublisherEscherRecord fdg = records.Reader.Records(records.Utility.Offset, records.Utility.End).Single(item => item.Kind == 0xF008);
            uint cluster = BitConverter.ToUInt32(art, fdg.Offset + 4) >> 10;
            uint used = invalid == 1 ? 1024U : 4U;
            Write(fdg.Offset + 4, (cluster << 10) + used - 1);
            Write(global.Offset + 16 + ((int)cluster - 1) * 8 + 4, used);
        } else Write(global.Offset + 8, BitConverter.ToUInt32(art, global.Offset + 8) + 1);
        byte[] input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(input);
        Assert.Equal(input, document.ToPubBytes());
        Apply(document, "TableBorderTopology");
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
        void Write(int offset, uint value) => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, art, offset, 4);
    }

    [Fact]
    public void Native_automatic_table_cell_default_is_white_and_remains_editable() {
        byte[] source = Fixture("TableBorderDefaultFill");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = document.Pages[0].Tables[0];
        Assert.Equal(OfficeColor.White, table.Cells[2].FillColor);
        Assert.Contains("#FFFFFF", document.ToSvg());
        Assert.Equal(source, document.ToPubBytes());
        document.SetTableCellFillColor(table.Id, 1, 0, OfficeColor.FromRgb(224, 192, 255));
        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.Equal(OfficeColor.FromRgb(224, 192, 255), PublisherDocument.Load(saved).Pages[0].Tables[0].Cells[2].FillColor);
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Escher/EscherStm");
    }

    internal static void Apply(PublisherDocument document, string variant) {
        PublisherTable table = document.Pages[0].Tables[0];
        uint id = table.Id;
        if (variant == "TableBordersCleared") {
            foreach (var edge in table.BorderSegments) document.SetTableBorderStyle(id, edge.FirstRow, edge.FirstColumn, edge.LastRow, edge.LastColumn, null);
            return;
        }
        document.SetTableBorderStyle(id, 0, 0, 1, 0, null);
        document.SetTableBorderStyle(id, 0, 0, 0, 2, new OfficeBorderSide(OfficeColor.FromRgb(0, 128, 0), 2));
        document.SetTableBorderStyle(id, 2, 0, 2, 1, new OfficeBorderSide(OfficeColor.FromRgb(255, 128, 0), 2));
        if (variant == "TableBorderDefaultFill") document.SetTableCellFillColor(id, 1, 0, OfficeColor.White);
        if (variant != "TableBorderTopologyTextFill") return;
        document.SetTableCellFillColor(id, 0, 0, null);
        document.SetTableCellFillColor(id, 1, 0, OfficeColor.FromRgb(224, 192, 255));
        OfficeRichTextParagraph template = table.Cells[0].Paragraphs[0]; OfficeRichTextRun run = template.Runs[0];
        document.ReplaceTableCellParagraphs(id, 0, 0, new[] { "Updated", "Item" }.Select(text => new OfficeRichTextParagraph(
            new[] { new OfficeRichTextRun(text, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily) },
            template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops)).ToArray());
    }

    internal static (int Row, int Column, bool Horizontal, OfficeBorderSide Side)[] Coverage(PublisherTable table) => table.BorderSegments.SelectMany(edge => {
        bool horizontal = edge.FirstRow == edge.LastRow;
        int start = horizontal ? edge.FirstColumn : edge.FirstRow, end = horizontal ? edge.LastColumn : edge.LastRow;
        return Enumerable.Range(start, end - start).Select(index => (Row: horizontal ? edge.FirstRow : index,
            Column: horizontal ? index : edge.FirstColumn, Horizontal: horizontal, Side: edge.Side));
    }).OrderBy(edge => edge.Horizontal).ThenBy(edge => edge.Row).ThenBy(edge => edge.Column).ToArray();

    private static (PublisherEscherReader Reader, PublisherEscherRecord Dgg, PublisherEscherRecord Utility) DrawingRecords(byte[] art) {
        var reader = new PublisherEscherReader(new PublisherBinaryData(art, "Escher/EscherStm"), null,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
        var top = reader.Records(0, art.Length);
        return (reader, top.Single(item => item.Kind == 0xF000), top.Single(item => item.Kind == 0xF002 && BitConverter.ToUInt32(art, item.Offset - 12) == 2));
    }
    private static (uint Cursor, uint Used, uint Maximum, uint Count, uint[] Ids) Allocation(byte[] source) {
        byte[] art = Compound(source).Streams["Escher/EscherStm"];
        var records = DrawingRecords(art);
        var fdg = records.Reader.Records(records.Utility.Offset, records.Utility.End).Single(item => item.Kind == 0xF008);
        var group = records.Reader.Records(records.Utility.Offset, records.Utility.End).Single(item => item.Kind == 0xF003);
        var global = records.Reader.Records(records.Dgg.Offset, records.Dgg.End).Single(item => item.Kind == 0xF006);
        uint cursor = BitConverter.ToUInt32(art, fdg.Offset + 4);
        uint[] ids = records.Reader.Records(group.Offset, group.End).Select(shape => records.Reader.Records(shape.Offset, shape.End).Single(item => item.Kind == 0xF00A))
            .Select(fsp => BitConverter.ToUInt32(art, fsp.Offset)).ToArray();
        Assert.Equal(ids.Length, ids.Distinct().Count());
        Assert.Equal((uint)ids.Length, BitConverter.ToUInt32(art, fdg.Offset));
        return (cursor, BitConverter.ToUInt32(art, global.Offset + 16 + ((int)(cursor >> 10) - 1) * 8 + 4),
            BitConverter.ToUInt32(art, global.Offset), BitConverter.ToUInt32(art, global.Offset + 8), ids);
    }
    private static byte[] Fixture(string variant) => File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + variant + "/native-" + variant.ToLowerInvariant() + ".pub"));
    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? result, out string? error), error); return result!;
    }
}
