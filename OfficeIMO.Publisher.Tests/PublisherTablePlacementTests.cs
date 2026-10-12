using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTablePlacementTests {
    [Theory]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    public void Native_table_data_survives_unavailable_placement(bool hidden, int geometry) {
        var source = new PublisherContents(100, 100, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        var grid = new PublisherSourceTable(new[] { 20D, 30D }, new[] { 10D });
        grid.Cells.Add(new PublisherSourceCell(0, 0, 0, 1));
        source.Shapes.Add(2, new PublisherSourceShape(new PublisherContentChunk(2, 0x10, 0, 1),
            new[] { new PublisherBlock(0x27, 0, 0, 0, 3) }) { Table = grid });
        source.ObjectPages.Add(2, 1);
        var text = new PublisherQuillText();
        var paragraphs = new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("retained", 10, OfficeColor.Black) }) };
        text.Stories.Add(3, new PublisherTextStory(3, "retained", paragraphs));
        text.CellParagraphs.Add(3, new List<IReadOnlyList<OfficeRichTextParagraph>> { paragraphs });
        var shape = new PublisherEscherShape(2, 1, 0, hidden ? new[] { new OfficeArtProperty(0, 0x03BF, 0x40004000) } : Array.Empty<OfficeArtProperty>()) {
            Bounds = geometry == 0 ? null : new PublisherNativeRectangle(0, 0, geometry == 2 ? 0 : 127000, 127000),
            GroupTransform = geometry == 1 ? OfficeTransform.Scale(0, 1) : OfficeTransform.Identity,
            LayoutTransform = geometry == 1 ? OfficeTransform.Scale(0, 1) : OfficeTransform.Identity
        };
        var escher = new PublisherEscherData(); escher.Shapes.Add(shape);
        PublisherDocument document = new PublisherPageProjector(source, text, escher,
            new PublisherParseContext(new PublisherReadOptions(), default)).Project();
        PublisherTable table = Assert.Single(document.Pages[0].Tables);
        Assert.False(table.HasPlacement);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TABLE_PLACEMENT_UNRESOLVED");
        Assert.True(table.HasTextMapping); Assert.Equal("retained", Assert.Single(table.Cells).Text);
        Assert.Equal(2, table.Cells[0].ColumnSpan);
        Assert.DoesNotContain("retained", document.ToSvg());
    }
}
