using OfficeIMO.Rtf;
using OfficeIMO.Html.Rtf;
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

        RtfDocument roundTrip = html.LoadFromHtml();
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

    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Hidden_Horizontal_Merge_Continuation_Cell_Metadata() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        RtfTable table = document.AddTable(1, 3);
        RtfTableCell first = table.Rows[0].Cells[0];
        RtfTableCell continuation = table.Rows[0].Cells[1];
        RtfTableCell last = table.Rows[0].Cells[2];
        first.HorizontalMerge = RtfTableCellMerge.First;
        continuation.HorizontalMerge = RtfTableCellMerge.Continue;
        first.RightBoundaryTwips = 1800;
        continuation.RightBoundaryTwips = 4200;
        continuation.SetPreferredWidth(2400, RtfTableWidthUnit.Twips).SetNoWrap().SetFitText();
        continuation.TopLeftToBottomRightBorder.Style = RtfTableCellBorderStyle.Dotted;
        continuation.TopLeftToBottomRightBorder.Width = 6;
        continuation.TopLeftToBottomRightBorder.ColorIndex = red;
        first.AddParagraph("Merged");
        last.AddParagraph("Tail");

        string html = document.ToHtml(new RtfHtmlSaveOptions { NewLine = "\n" });

        Assert.Contains("colspan=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-cell=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(roundTrip.Blocks));
        RtfTableCell roundTripFirst = roundTripTable.Rows[0].Cells[0];
        RtfTableCell roundTripContinuation = roundTripTable.Rows[0].Cells[1];
        Assert.Equal(RtfTableCellMerge.First, roundTripFirst.HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, roundTripContinuation.HorizontalMerge);
        Assert.Equal(1800, roundTripFirst.RightBoundaryTwips);
        Assert.Equal(4200, roundTripContinuation.RightBoundaryTwips);
        Assert.Equal(2400, roundTripContinuation.PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Twips, roundTripContinuation.PreferredWidthUnit);
        Assert.True(roundTripContinuation.NoWrap);
        Assert.True(roundTripContinuation.FitText);
        Assert.Equal(RtfTableCellBorderStyle.Dotted, roundTripContinuation.TopLeftToBottomRightBorder.Style);
        Assert.Equal(6, roundTripContinuation.TopLeftToBottomRightBorder.Width);
        Assert.Equal(red, roundTripContinuation.TopLeftToBottomRightBorder.ColorIndex);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\clmgf\cellx1800", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clmrg\clftsWidth3\clwWidth2400\clNoWrap\clFitText\cldglu\brdrdot\brdrw6\brdrcf1\cellx4200", rtf, StringComparison.Ordinal);
    }
}
