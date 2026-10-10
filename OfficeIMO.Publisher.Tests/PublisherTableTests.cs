using OfficeIMO.Drawing;
using OfficeIMO.TestSupport;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableTests {
    [Theory]
    [InlineData("Sample.pub")]
    [InlineData("Sample_2010.pub")]
    public void Native_table_exposes_tracks_cells_styled_text_and_page_placement(string fixture) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(fixture));
        PublisherTable table = Assert.Single(document.Pages[1].Tables);
        Assert.Equal(299U, table.Id); Assert.Equal(document.Pages[1].Id, table.PageId);
        Assert.True(table.HasTextMapping); Assert.NotNull(table.StoryId);
        Assert.Equal(3, table.RowCount); Assert.Equal(2, table.ColumnCount);
        Assert.Equal(new[] { "Table on page 2", "Top right", "P2 table left", "P2 table right", "Bottom Left", "Bottom Right" }, table.Cells.Select(cell => cell.Text));
        Assert.Equal(new[] { 0, 0, 1, 1, 2, 2 }, table.Cells.Select(cell => cell.RowIndex));
        Assert.Equal(new[] { 0, 1, 0, 1, 0, 1 }, table.Cells.Select(cell => cell.ColumnIndex));
        Assert.All(table.Cells, cell => { Assert.Equal(1, cell.RowSpan); Assert.Equal(1, cell.ColumnSpan); Assert.NotEmpty(cell.Paragraphs); });
        Assert.Equal(table.ColumnWidths[0], table.Cells[0].Width, 6);
        Assert.Equal(table.RowHeights[0], table.Cells[0].Height, 6);
        Assert.Equal(table.Cells[0].Width, table.Cells[1].X, 6);
        Assert.Equal(table.Cells[0].Height, table.Cells[2].Y, 6);
        OfficePoint origin = table.PageTransform.TransformPoint(new OfficePoint(0, 0));
        Assert.Equal(table.X, origin.X, 6); Assert.Equal(table.Y, origin.Y, 6);
        Assert.Empty(document.Pages[0].Tables); Assert.All(document.MasterPages, page => Assert.Empty(page.Tables));
    }

    [Fact]
    public void Merged_cell_retains_its_span_geometry_and_both_native_text_regions() {
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture("Sample.pub"));
        PublisherDocument document = PublisherDocument.Load(PublisherTableFixture.MergeFirstRow(original));
        PublisherTable table = Assert.Single(document.Pages[1].Tables);
        PublisherTableCell merged = table.Cells[0];
        Assert.Equal(5, table.Cells.Count); Assert.True(table.HasTextMapping);
        Assert.Equal(0, merged.RowIndex); Assert.Equal(0, merged.ColumnIndex);
        Assert.Equal(1, merged.RowSpan); Assert.Equal(2, merged.ColumnSpan);
        Assert.Equal(table.ColumnWidths.Sum(), merged.Width, 6);
        Assert.Equal("Table on page 2\nTop right", merged.Text);
        Assert.Equal(PublisherDocument.Load(original).TextStories.Select(story => story.Text), document.TextStories.Select(story => story.Text));
    }

    [Fact]
    public void Unresolved_text_mapping_preserves_the_grid_and_complete_story_with_an_omission() {
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture("Sample.pub"));
        PublisherDocument document = PublisherDocument.Load(PublisherTableFixture.WithoutTextMapping(original));
        PublisherTable table = Assert.Single(document.Pages[1].Tables);
        Assert.False(table.HasTextMapping); Assert.Equal(6, table.Cells.Count);
        Assert.All(table.Cells, cell => { Assert.Empty(cell.Paragraphs); Assert.Empty(cell.Text); });
        Assert.Equal(PublisherDocument.Load(original).TextStories.Select(story => story.Text), document.TextStories.Select(story => story.Text));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TABLE_TEXT_MAPPING_UNRESOLVED"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void Master_table_is_owned_once_while_its_artwork_is_inherited_by_document_pages() {
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture("Sample.pub"));
        PublisherDocument baseline = PublisherDocument.Load(original);
        PublisherPage master = Assert.Single(baseline.MasterPages);
        PublisherDocument document = PublisherDocument.Load(PublisherTableFixture.MoveToMaster(original, baseline.Pages[1].Id, master.Id));
        Assert.All(document.Pages, page => Assert.Empty(page.Tables));
        PublisherTable table = Assert.Single(Assert.Single(document.MasterPages).Tables);
        Assert.Equal(master.Id, table.PageId); Assert.True(table.HasTextMapping);
        Assert.All(document.Pages, page => Assert.Contains("Table on page 2", document.ToSvg(document.Pages.ToList().IndexOf(page))));
    }
}
