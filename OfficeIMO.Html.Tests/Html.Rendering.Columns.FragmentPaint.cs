using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData(true, HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData(false, HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    [InlineData(true, HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void HtmlColumns_FragmentPaintOmitsRuleBesideASingleTerminalColumn(
        bool descendant, HtmlRenderIntentProfile profile) {
        string content = DescendantSpannerParagraphs("P", 4)
            + "<h2 class='spanner'>SPAN</h2>" + DescendantSpannerParagraphs("T", 6);
        HtmlRenderDocument rendered = RenderDescendantSpanners(
            descendant ? "<div>" + content + "</div>" : content, profile,
            extraCss: "#columns{column-rule:2px solid #0066cc}");

        AssertDescendantSpannerSourceOnce(rendered,
            "P1", "P2", "P3", "P4", "SPAN", "T1", "T2", "T3", "T4", "T5", "T6", "AFTER");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        // Edge N01 paints the balanced pair's rule, but none beside the final
        // single occupied column. The requested column count remains two.
        Assert.Equal(OfficeColor.FromRgb(0, 102, 204), raster.GetPixel(149, 10));
        Assert.Equal(OfficeColor.White, raster.GetPixel(149, 100));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_FragmentPaintKeepsRuleThroughASpannerKeepDeferral(HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<div>"
            + DescendantSpannerParagraphs("P", 2)
            + "<div class='spanner' style='break-inside:avoid-page'>"
            + DescendantSpannerParagraphs("S", 3) + "</div>"
            + DescendantSpannerParagraphs("T", 8) + "</div>", profile,
            extraCss: "#columns{column-rule:2px solid #0066cc}", entry: "<p>E1</p><p>E2</p>");

        Assert.Equal(2, DescendantSpannerTextPage(rendered, "S1").PageNumber);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        // Edge N15 continues the preceding set's rule through the unused body
        // tail while the kept spanner starts on the next physical page.
        Assert.Equal(OfficeColor.FromRgb(0, 102, 204), raster.GetPixel(169, 100));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(false, HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(true, HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(false, HtmlRenderIntentProfile.ScreenMediaPaged)]
    [InlineData(true, HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_FragmentPaintContinuesSliceBackdropAndSidesToPhysicalCut(
        bool emptyLeading, HtmlRenderIntentProfile profile) {
        string content = emptyLeading
            ? "<p>P1</p><p>P2</p><div id='wrapper'><h2 class='spanner'>SPAN</h2>"
                + DescendantSpannerParagraphs("T", 4) + "</div>"
            : "<div id='wrapper'>" + DescendantSpannerParagraphs("P", 4)
                + "<h2 class='spanner'>SPAN</h2>" + DescendantSpannerParagraphs("T", 6) + "</div>";
        string decoration = emptyLeading
            ? "margin-top:6px;border:2px solid #008833;padding-top:8px;background:#d9f0d9"
            : "margin-left:10%;width:80%;padding:8px 5%;border:2px solid #008833;background:#d9f0d9;box-decoration-break:slice";
        HtmlRenderDocument rendered = RenderDescendantSpanners(content, profile,
            extraCss: "#columns{column-rule:2px solid #0066cc}#wrapper{" + decoration + "}");

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        // Edge N04/N06 preserve source line cuts, but paint the continuing
        // ancestor through the remaining 10px/4px of this physical band.
        Assert.Equal(OfficeColor.FromRgb(217, 240, 217), raster.GetPixel(100, emptyLeading ? 118 : 115));
        Assert.Equal(OfficeColor.FromRgb(0, 136, 51), raster.GetPixel(emptyLeading ? 21 : 163, emptyLeading ? 118 : 115));
        Assert.Equal(emptyLeading ? 2 : 3, DescendantSpannerTextPage(rendered, "AFTER").PageNumber);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlColumns_FragmentPaintFiniteOverflowHonorsOriginatingColumnReservations(bool note) {
        string first = note
            ? "<p>CALL<span style='float:footnote;float-reference:column;font:14px/20px Arial'>RESERVED</span></p>"
            : "<p>CALL</p><aside style='float:bottom;float-reference:column;height:20px'>RESERVED</aside>";
        HtmlRenderDocument rendered = RenderDescendantSpanners(first
            + "<div style='height:20px;overflow:visible'>" + DescendantSpannerParagraphs("C", 4) + "</div>",
            HtmlRenderIntentProfile.PrintPaged, extraCss: "#columns{height:100px}");

        HtmlRenderText reserved = DescendantSpannerText(rendered, "RESERVED");
        foreach (int index in Enumerable.Range(1, 4)) {
            HtmlRenderText text = DescendantSpannerText(rendered, "C" + index);
            Assert.Equal(1, DescendantSpannerTextPage(rendered, text.Text).PageNumber);
            if (text.X < 160D)
                Assert.True(text.Y + text.LineHeight <= reserved.Y + 0.001D,
                    "Finite overflow must stop before the existing note or bottom figure reservation.");
        }
        Assert.Equal(180D, DescendantSpannerText(rendered, note ? "C3" : "C4").X, 3);
        AssertColumnBreaksQualified(rendered);
    }
}
