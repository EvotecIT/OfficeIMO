using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableBorderExtentWritingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Native_save_refuses_borders_that_no_longer_fit_after_table_resize(bool existingUtilities, bool vertical) {
        string fixture = existingUtilities ? "TableDecoration/native-tabledecoration.pub" : "GeneratedTable/generated-table.pub";
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherTable table = document.Pages[0].Tables[0];
        document.SetTableBorderStyle(table.Id, 0, 0, vertical ? 1 : 0, vertical ? 0 : 1, new OfficeBorderSide(OfficeColor.Black, 24));
        if (existingUtilities) {
            // Exercise a bounds-only save of a previously valid wide edge,
            // independently of the first-utility allocation branch.
            document = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
            table = document.Pages[0].Tables[0];
        }
        document.SetTableBounds(table.Id, table.X, table.Y, vertical ? 15 : table.Width, vertical ? table.Height : 12);
        Assert.Contains(document.Pages[0].Tables[0].BorderSegments, edge => edge.Side.Width == 24);
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Fresh_utility_encoder_refuses_an_authored_edge_wider_than_the_final_table_extent(bool vertical) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        uint id = document.AddTable(document.Pages[0].Id, 72, 144, new[] { 120D, 180D }, new[] { 42D, 54D }).Id;
        document.SetTableBorderStyle(id, 0, 0, vertical ? 1 : 0, vertical ? 0 : 1, new OfficeBorderSide(OfficeColor.Black, 24));
        document.SetTableBounds(id, 72, 144, vertical ? 15 : 300, vertical ? 96 : 12);
        PublisherTable table = document.Pages[0].Tables[0];
        // The source-free specimen consumes this codec directly; public native
        // creation remains gated before reaching it.
        var input = new PublisherTableUtilityInput(id, table.RowHeights.Count, table.ColumnWidths.Count, table.Width, table.Height,
            table.Cells.Select(cell => new PublisherTableCellSlot(cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan)).ToArray(),
            table.Cells.Select(cell => cell.FillColor).ToArray(), table.BorderSegments);
        Assert.Null(PublisherTableUtilityDrawingEncoding.Write(3, new[] { input }, 1024, CancellationToken.None));
    }
}
