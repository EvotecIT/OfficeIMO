using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlTableCellMetadataTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Table_Cell_Rtf_Metadata() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        int blue = document.AddColor(0, 0, 255);
        RtfTable table = document.AddTable(1, 2);
        RtfTableCell first = table.Rows[0].Cells[0];
        RtfTableCell second = table.Rows[0].Cells[1];
        first.RightBoundaryTwips = 1800;
        second.RightBoundaryTwips = 5000;
        first.TopLeftToBottomRightBorder.Style = RtfTableCellBorderStyle.Dotted;
        first.TopLeftToBottomRightBorder.Width = 6;
        first.TopLeftToBottomRightBorder.ColorIndex = red;
        first.TopRightToBottomLeftBorder.Style = RtfTableCellBorderStyle.Dashed;
        first.TopRightToBottomLeftBorder.Width = 10;
        first.TopRightToBottomLeftBorder.ColorIndex = blue;
        first.AddParagraph("Left");
        second.AddParagraph("Right");

        string html = document.ToHtml(new RtfHtmlSaveOptions { NewLine = "\n" });

        Assert.Contains("data-officeimo-rtf-cell=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.ToRtfDocumentFromHtml();
        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(roundTrip.Blocks));
        RtfTableCell roundTripFirst = roundTripTable.Rows[0].Cells[0];
        RtfTableCell roundTripSecond = roundTripTable.Rows[0].Cells[1];
        Assert.Equal(1800, roundTripFirst.RightBoundaryTwips);
        Assert.Equal(5000, roundTripSecond.RightBoundaryTwips);
        Assert.Equal(RtfTableCellBorderStyle.Dotted, roundTripFirst.TopLeftToBottomRightBorder.Style);
        Assert.Equal(6, roundTripFirst.TopLeftToBottomRightBorder.Width);
        Assert.Equal(red, roundTripFirst.TopLeftToBottomRightBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, roundTripFirst.TopRightToBottomLeftBorder.Style);
        Assert.Equal(10, roundTripFirst.TopRightToBottomLeftBorder.Width);
        Assert.Equal(blue, roundTripFirst.TopRightToBottomLeftBorder.ColorIndex);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\cldglu\brdrdot\brdrw6\brdrcf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cldgll\brdrdash\brdrw10\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cellx1800", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cellx5000", rtf, StringComparison.Ordinal);
    }
}
