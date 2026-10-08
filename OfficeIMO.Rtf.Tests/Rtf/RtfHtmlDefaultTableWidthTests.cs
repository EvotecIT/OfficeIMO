using OfficeIMO.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Rtf.Tests;

public class RtfHtmlDefaultTableWidthTests {
    [Theory]
    [InlineData(2, 4800)]
    [InlineData(5, 8640)]
    public void OrdinaryHtmlDefaultColumnsFitThePageAfterSaveAndReopen(int columns, int expectedRightEdge) {
        string cells = string.Concat(Enumerable.Range(1, columns).Select(i => "<td>Column " + i + "</td>"));
        RtfDocument document = HtmlConversionDocument.Parse("<table><tr>" + cells + "</tr></table>").ToRtfDocument();
        RtfDocument reopened = RtfDocument.Read(document.ToRtf()).Document;
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(reopened.Blocks));
        Assert.Equal(expectedRightEdge, table.Rows[0].Cells.Last().RightBoundaryTwips);
        Assert.Equal(columns, table.Rows[0].Cells.Count);
        Assert.Equal("Column " + columns, table.Rows[0].Cells.Last().Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void OrdinaryNestedTableUsesTheFittedHostCell() {
        string cells = string.Concat(Enumerable.Range(1, 5).Select(i => "<td>Column " + i + "</td>"));
        RtfDocument document = HtmlConversionDocument.Parse("<table><tr><td>Outer</td><td><table><tr>" + cells + "</tr></table></td></tr></table>").ToRtfDocument();
        RtfTable outer = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        RtfTable inner = Assert.IsType<RtfTable>(Assert.Single(outer.Rows[0].Cells[1].Blocks));
        Assert.InRange(inner.Rows[0].Cells.Last().RightBoundaryTwips!.Value, 1, 2400);
    }

    [Fact]
    public void RoundTripTableWithAuthoredBoundariesKeepsItsGeometry() {
        RtfDocument original = RtfDocument.Create();
        RtfTable table = original.AddTable(1, 5);
        for (int i = 0; i < 5; i++) table.Rows[0].Cells[i].AddParagraph("Column " + i);
        string html = original.ToHtml(new RtfToHtmlOptions { IncludeRoundTripMetadata = true });
        RtfDocument reopened = HtmlConversionDocument.Parse(html).ToRtfDocument();
        RtfTable restored = Assert.IsType<RtfTable>(Assert.Single(reopened.Blocks));
        Assert.Equal(12000, restored.Rows[0].Cells.Last().RightBoundaryTwips);
    }
    [Fact]
    public void SparseRowWithTrailingRowspanFitsLogicalColumnBoundaries() {
        string html = "<table><tr><td>A</td><td>B</td><td>C</td><td>D</td><td rowspan='2'>E</td></tr><tr><td>F</td></tr></table>";
        RtfDocument document = HtmlConversionDocument.Parse(html).ToRtfDocument();
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        Assert.Equal(8640, table.Rows[0].Cells.Last().RightBoundaryTwips);
        Assert.Equal(8640, table.Rows[1].Cells.Last().RightBoundaryTwips);
        Assert.Equal(RtfTableCellMerge.Continue, table.Rows[1].Cells.Last().VerticalMerge);
        Assert.Equal("F", table.Rows[1].Cells[0].Paragraphs[0].ToPlainText());
    }

    [Theory]
    [InlineData("", 1512)]
    [InlineData("width='120'", 1800)]
    public void TableFittingRefitsOnlyIntrinsicImages(string attributes, int expectedWidth) {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(120, 60, OfficeColor.Red));
        string image = "<img " + attributes + " src='data:image/png;base64," + Convert.ToBase64String(png) + "'>";
        string html = "<table><tr><td>A</td><td>B</td><td>C</td><td>D</td><td>" + image + "</td></tr></table>";
        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        RtfDocument reopened = RtfDocument.Read(result.RequireValue().ToRtf()).Document;
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(reopened.Blocks));
        RtfImage photo = Assert.Single(table.Rows[0].Cells[4].Paragraphs.SelectMany(p => p.Inlines).OfType<RtfImage>());
        Assert.Equal(png, photo.Data);
        Assert.Equal(expectedWidth, photo.DesiredWidthTwips);
        if (attributes.Length == 0) {
            Assert.Equal(756, photo.DesiredHeightTwips);
            Assert.Contains(result.RtfDiagnostics, diagnostic => diagnostic.Code == "HtmlRtfImageFittedToCell");
        }
    }

}
