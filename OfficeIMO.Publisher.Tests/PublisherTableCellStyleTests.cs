using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableCellStyleTests {
    [Theory]
    [InlineData("TableDecoration/native-tabledecoration.pub", false)]
    [InlineData("TableDecorationEdited/native-tabledecorationedited.pub", true)]
    public void Independent_native_cell_paint_keeps_definition_ownership_and_shared_grid_edges(string fixture, bool edited) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        Check(document, edited);
        PublisherPage page = Assert.Single(document.Pages);
        Assert.Equal(2, page.ObjectIds.Count);
        Assert.Single(page.TextFrames);
        Assert.Equal(new[] { page.TextFrames[0].Id }, page.Shapes.Select(shape => shape.Id));
        Assert.Equal(new[] { "Item", "Count", "Alpha", "" }, page.Tables[0].Cells.Select(cell => cell.Text));
        string svg = document.ToSvg();
        Assert.Contains(edited ? "#C0E0C0" : "#FFE0A0", svg, StringComparison.OrdinalIgnoreCase);
        foreach (string color in new[] { "#A0E0FF", "#B4F0B4", "#0000FF", "#008000", "#800080", "#FF00FF", "#00FFFF" })
            Assert.Contains(color, svg, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("TableDecoration/native-tabledecoration.pub", false)]
    [InlineData("TableDecorationEdited/native-tabledecorationedited.pub", true)]
    public void Cell_text_growth_preserves_native_decoration_and_opaque_artwork(string fixture, bool edited) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = document.Pages[0].Tables[0];
        OfficeRichTextParagraph template = original.Cells[0].Paragraphs[0];
        OfficeRichTextRun run = template.Runs[0];
        var paragraphs = new[] { "Updated", "Item" }.Select(text => new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun(text, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily)
        }, template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops)).ToArray();
        document.ReplaceTableCellParagraphs(original.Id, 0, 0, paragraphs);
        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Quill/QuillSub/CONTENTS");
        PublisherDocument reopened = PublisherDocument.Load(saved);
        Check(reopened, edited);
        Assert.Equal("Updated\nItem", reopened.Pages[0].Tables[0].Cells[0].Text);
        Assert.Equal("Item", original.Cells[0].Text);
        Assert.Equal(new[] { "Count", "Alpha", "" }, reopened.Pages[0].Tables[0].Cells.Skip(1).Select(cell => cell.Text));
    }

    [Fact]
    public void Borderless_native_control_has_no_recovered_cell_fill_or_edge() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TableCreation/native-tablecreation.pub"));
        PublisherTable table = Assert.Single(document.Pages[0].Tables);
        Assert.All(table.Cells, cell => Assert.Null(cell.FillColor));
        Assert.Empty(table.BorderSegments);
    }

    [Fact]
    public void Unsupported_native_cell_fill_is_reported_without_losing_opaque_save_bytes() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableDecoration/native-tabledecoration.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        byte[] art = (byte[])compound!.Streams["Escher/EscherStm"].Clone();
        // Replace the uniquely colored solid-fill property with an unsupported
        // gradient selector inside the same six-byte native property entry.
        byte[] entry = { 0x81, 0x01, 0xFF, 0xE0, 0xA0, 0x00 };
        int offset = Assert.Single(Enumerable.Range(0, art.Length - entry.Length + 1),
            index => entry.Select((value, delta) => art[index + delta] == value).All(value => value));
        art[offset] = 0x80;
        Buffer.BlockCopy(BitConverter.GetBytes(4U), 0, art, offset + 2, 4);
        byte[] input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherTable table = document.Pages[0].Tables[0];
        Assert.Null(table.Cells[0].FillColor);
        Assert.Equal(OfficeColor.FromRgb(160, 224, 255), table.Cells[1].FillColor);
        Assert.Equal(6, table.BorderSegments.Count);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TABLE_CELL_STYLE_UNSUPPORTED"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(input, document.ToPubBytes());
        document.SetTableCellFillColor(table.Id, 0, 0, OfficeColor.FromRgb(20, 40, 60));
        Assert.Throws<OfficeIMO.OfficeConversionException>(() => document.ToPubBytes(
            new PublisherSaveOptions { LossPolicy = OfficeIMO.OfficeConversionLossPolicy.Allow }));
    }

    [Theory]
    [InlineData(0U)]
    [InlineData(293U)]
    [InlineData(uint.MaxValue)]
    public void Unresolved_cell_artwork_owner_reports_omission_and_preserves_source(uint owner) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableDecoration/native-tabledecoration.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        byte[] art = (byte[])compound!.Streams["Escher/EscherStm"].Clone();
        // This control's first table-owned utility anchor is the first cell fill.
        byte[] anchor = { 0x02, 0x68, 0x26, 0x01, 0x00, 0x00 };
        int offset = Enumerable.Range(0, art.Length - anchor.Length + 1)
            .First(index => anchor.Select((value, delta) => art[index + delta] == value).All(value => value));
        Buffer.BlockCopy(BitConverter.GetBytes(owner), 0, art, offset + 2, 4);
        byte[] input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = art });
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherTable table = document.Pages[0].Tables[0];
        Assert.Null(table.Cells[0].FillColor);
        Assert.Equal(OfficeColor.FromRgb(160, 224, 255), table.Cells[1].FillColor);
        Assert.Equal(6, table.BorderSegments.Count);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TABLE_CELL_STYLE_UNSUPPORTED"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(input, document.ToPubBytes());
    }

    private static void Check(PublisherDocument document, bool edited) {
        PublisherTable table = Assert.Single(document.Pages[0].Tables);
        Assert.Equal(new[] { 120D, 180D }, table.ColumnWidths);
        Assert.Equal(new[] { 42D, 54D }, table.RowHeights);
        Assert.Equal(new OfficeColor?[] {
            edited ? OfficeColor.FromRgb(192, 224, 192) : OfficeColor.FromRgb(255, 224, 160),
            OfficeColor.FromRgb(160, 224, 255), null, OfficeColor.FromRgb(180, 240, 180)
        }, table.Cells.Select(cell => cell.FillColor));
        var expected = new[] {
            (0,0,1,0, new OfficeBorderSide(edited ? OfficeColor.FromRgb(0,128,0) : OfficeColor.FromRgb(255,0,0), edited ? 1 : 2)),
            (0,0,0,1, new OfficeBorderSide(OfficeColor.FromRgb(0,0,255), 3)),
            (0,1,1,1, new OfficeBorderSide(OfficeColor.FromRgb(0,128,0), 1)),
            (1,0,1,1, new OfficeBorderSide(OfficeColor.FromRgb(128,0,128), edited ? 3 : 1)),
            (1,2,2,2, new OfficeBorderSide(OfficeColor.FromRgb(255,0,255), 4)),
            (2,1,2,2, new OfficeBorderSide(OfficeColor.FromRgb(0,255,255), 2))
        };
        Assert.Equal(expected, table.BorderSegments.Select(border => (border.FirstRow, border.FirstColumn, border.LastRow, border.LastColumn, border.Side)));
    }
}
