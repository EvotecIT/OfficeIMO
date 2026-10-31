using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableTrackReadingTests {
    [Theory]
    [InlineData("TableLabelShortParagraphs", "native-tablelabelshortparagraphs.pub", 457200)]
    [InlineData("TableLabelWrapped", "native-tablelabelwrapped.pub", 777247)]
    public void Native_table_cells_use_saved_layout_endpoints_instead_of_nominal_row_sizes(string folder, string leaf, int firstRowEmu) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture($"NativeComparison/{folder}/{leaf}"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = Assert.Single(document.Pages[0].Tables);
        double firstRow = firstRowEmu / 12700D;
        Assert.Equal(new[] { 90D, 90D }, table.ColumnWidths);
        Assert.Equal(new[] { firstRow, 36D }, table.RowHeights);
        Assert.Equal(table.RowHeights.Sum(), table.Height, 9);
        Assert.Equal(firstRow, table.Cells[0].Height, 9);
        Assert.Equal(firstRow, table.Cells[2].Y, 9);
        Assert.Equal(90D, table.Cells[1].X);
        Assert.Equal(36D, table.Cells[2].Height, 9);
        Assert.Equal("Alpha", table.Cells[2].Text);
        Assert.Equal(2, table.Cells[0].Paragraphs.Count);
        Assert.Equal(source, document.ToPubBytes());
    }

    [Fact]
    public void Missing_saved_track_coordinates_retain_nominal_geometry_with_an_explicit_report() {
        byte[] source = MutateTrack((bytes, coordinate) => bytes[coordinate.Offset - 2] = 7, 2);
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = Assert.Single(document.Pages[0].Tables);
        Assert.Equal(new[] { 36D, 36D }, table.RowHeights);
        Assert.Equal(new[] { 90D, 90D }, table.ColumnWidths);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, finding => finding.Code == "PUB_TABLE_TRACK_POSITIONS_UNASSESSED"
            && finding.LossKind == OfficeConversionLossKind.Unassessed && finding.Location == "Contents/object/294");
        Assert.Equal(source, document.ToPubBytes());
    }

    [Fact]
    public void Nonincreasing_saved_track_coordinates_are_rejected_before_cell_projection() {
        byte[] source = MutateTrack((bytes, coordinate) => BitConverter.GetBytes(777247U).CopyTo(bytes, coordinate.Offset), 3);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(source));
        Assert.Contains("track", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] MutateTrack(Action<byte[], PublisherBlock> mutation, int index) =>
        PublisherInputContractTests.Mutate("Contents", bytes => {
            var data = new PublisherBinaryData(bytes, "Contents");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            PublisherContents contents = new PublisherContentsReader(data, context).Read();
            PublisherSourceShape table = contents.Shapes.Values.Single(shape => shape.Table != null);
            var blocks = new PublisherBlockReader(data, context);
            PublisherBlock dimensions = table.Fields.Single(field => field.Id == 0x6D);
            PublisherBlock track = blocks.Children(dimensions).Where(field => field.Id == 0 && field.Type == 0x88).ElementAt(index);
            mutation(bytes, blocks.Children(track).Single(field => field.Id == 1));
        }, "NativeComparison/TableLabelWrapped/native-tablelabelwrapped.pub");
}
