using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlTableRowMetadataTests {
    [Fact]
    public void HtmlTableRowsStayTogetherInRtfUnlessRoundTripMetadataOverridesThem() {
        RtfDocument imported = HtmlConversionDocument.Parse(
            "<table><tr><th>Term</th><th>Definition</th></tr>"
            + "<tr><td>Basic</td><td>A wrapped definition</td></tr></table>").ToRtfDocument();
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(imported.Blocks));

        Assert.All(table.Rows, row => Assert.True(row.KeepTogether));
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(
            imported.ToRtf(new RtfWriteOptions { IncludeGenerator = false }), @"\\trkeep(?!follow)").Count);

        RtfDocument source = RtfDocument.Create();
        source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Splittable");
        string html = source.ToHtml(new RtfToHtmlOptions { IncludeRoundTripMetadata = true });
        RtfTable roundTrip = Assert.IsType<RtfTable>(Assert.Single(HtmlConversionDocument.Parse(html).ToRtfDocument().Blocks));
        Assert.False(Assert.Single(roundTrip.Rows).KeepTogether);
    }

    [Fact]
    public void OlderRoundTripRowMetadataWithoutKeepTogetherPreservesSplittableRow() {
        string legacyMetadata = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(
            "row.keepWithNext=dHJ1ZQ==\n"));
        string html = $"<table><tr data-officeimo-rtf-row=\"{legacyMetadata}\"><td>Legacy row</td></tr></table>";

        RtfDocument imported = HtmlConversionDocument.Parse(html).ToRtfDocument();
        RtfTableRow row = Assert.Single(Assert.IsType<RtfTable>(Assert.Single(imported.Blocks)).Rows);

        Assert.False(row.KeepTogether);
        Assert.True(row.KeepWithNext);
    }

    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Table_Row_Rtf_Metadata() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        int blue = document.AddColor(0, 0, 255);
        RtfTable table = document.AddTable(1, 1);
        RtfTableRow row = table.Rows[0];
        row.KeepTogether = true;
        row.KeepWithNext = true;
        row.SetAutoFit(false)
            .SetCellGap(240)
            .SetLeftIndent(720)
            .SetSpacing(topTwips: 10, leftTwips: 20, bottomTwips: 30, rightTwips: 40)
            .SetPositionAnchors(RtfTableHorizontalAnchor.Margin, RtfTableVerticalAnchor.Page)
            .SetPosition(RtfTableHorizontalPosition.Center, null, RtfTableVerticalPosition.Bottom, null)
            .SetTextWrapDistances(leftTwips: 187, rightTwips: 188, topTwips: 189, bottomTwips: 190);
        row.NoOverlap = true;
        row.TopBorder.Style = RtfTableCellBorderStyle.Single;
        row.TopBorder.Width = 12;
        row.TopBorder.ColorIndex = red;
        row.VerticalBorder.Style = RtfTableCellBorderStyle.Double;
        row.VerticalBorder.Width = 8;
        row.VerticalBorder.ColorIndex = blue;
        row.Cells[0].AddParagraph("Floating");

        string html = document.ToHtml(new RtfToHtmlOptions {
            IncludeRoundTripMetadata = true,
            EmbedImagesAsDataUri = true, NewLine = "\n" });

        Assert.Contains("data-officeimo-rtf-row=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = HtmlConversionDocument.Parse(html).ToRtfDocument();
        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(roundTrip.Blocks));
        RtfTableRow roundTripRow = Assert.Single(roundTripTable.Rows);
        Assert.True(roundTripRow.KeepTogether);
        Assert.True(roundTripRow.KeepWithNext);
        Assert.Equal(false, roundTripRow.AutoFit);
        Assert.Equal(240, roundTripRow.CellGapTwips);
        Assert.Equal(720, roundTripRow.LeftIndentTwips);
        Assert.Equal(10, roundTripRow.SpacingTopTwips);
        Assert.Equal(20, roundTripRow.SpacingLeftTwips);
        Assert.Equal(30, roundTripRow.SpacingBottomTwips);
        Assert.Equal(40, roundTripRow.SpacingRightTwips);
        Assert.True(roundTripRow.NoOverlap);
        Assert.Equal(RtfTableHorizontalAnchor.Margin, roundTripRow.HorizontalAnchor);
        Assert.Equal(RtfTableVerticalAnchor.Page, roundTripRow.VerticalAnchor);
        Assert.Equal(RtfTableHorizontalPosition.Center, roundTripRow.HorizontalPosition);
        Assert.Null(roundTripRow.HorizontalPositionTwips);
        Assert.Equal(RtfTableVerticalPosition.Bottom, roundTripRow.VerticalPosition);
        Assert.Null(roundTripRow.VerticalPositionTwips);
        Assert.Equal(187, roundTripRow.TextWrapLeftTwips);
        Assert.Equal(188, roundTripRow.TextWrapRightTwips);
        Assert.Equal(189, roundTripRow.TextWrapTopTwips);
        Assert.Equal(190, roundTripRow.TextWrapBottomTwips);
        Assert.Equal(RtfTableCellBorderStyle.Single, roundTripRow.TopBorder.Style);
        Assert.Equal(12, roundTripRow.TopBorder.Width);
        Assert.Equal(red, roundTripRow.TopBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Double, roundTripRow.VerticalBorder.Style);
        Assert.Equal(8, roundTripRow.VerticalBorder.Width);
        Assert.Equal(blue, roundTripRow.VerticalBorder.ColorIndex);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\trowd\trgaph240\trkeep\trkeepfollow\trautofit0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trspdt10\trspdft3\trspdl20\trspdfl3\trspdb30\trspdfb3\trspdr40\trspdfr3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tabsnoovrlp\tphmrg\tpvpg\tposxc\tposyb\tdfrmtxtLeft187\tdfrmtxtRight188\tdfrmtxtTop189\tdfrmtxtBottom190", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbrdrt\brdrs\brdrw12\brdrcf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbrdrv\brdrdb\brdrw8\brdrcf2", rtf, StringComparison.Ordinal);
    }
}

internal static class RtfHtmlTableTestMarkup {
    internal static string WithoutRowMetadata(string html) =>
        System.Text.RegularExpressions.Regex.Replace(
            html, " data-officeimo-rtf-row=\"[^\"]*\"", string.Empty);
}
