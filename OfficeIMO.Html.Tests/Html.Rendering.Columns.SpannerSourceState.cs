using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void HtmlColumns_SpannerSourceStatePaintsOneAbsoluteChildFromEitherAncestorSelection(bool beforeSpanner) {
        HtmlRenderDocument rendered = RenderDescendantSpanners(DescendantSpannerPositionedContent(beforeSpanner));

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "ABS", "B1", "B2", "SPAN", "C1", "C2", "T1", "T2", "T3", "T4", "AFTER");
        HtmlRenderText absolute = DescendantSpannerText(rendered, "ABS");
        // Native N19/N20 place the wrapper's initial fragment in column two:
        // its 160px origin plus the authored 80px absolute offset gives 240px.
        Assert.Equal(240D, absolute.X, 3);
        Assert.Equal(0D, absolute.Y, 3);
        Assert.Equal(160D, DescendantSpannerText(rendered, "B1").X, 3);
        Assert.Equal(300D, DescendantSpannerBackground(rendered, "h2#spanner").Width, 3);
        Assert.Equal(180D, DescendantSpannerText(rendered, "AFTER").Y, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(true, HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(false, HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(true, HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData(false, HtmlRenderIntentProfile.ScreenFullPage)]
    public void HtmlColumns_SpannerSourceStateAbsolutePaintDoesNotReorderFollowingInFlowText(
        bool beforeSpanner, HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = RenderDescendantSpanners(DescendantSpannerPositionedContent(beforeSpanner), profile);

        string[] labels = { "P1", "P2", "B1", "B2", "SPAN", "C1", "C2", "T1", "T2", "T3", "T4", "AFTER" };
        AssertDescendantSpannerSourceOnce(rendered, labels.Append("ABS").ToArray());
        Assert.Equal(labels, rendered.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Where(word => word != "ABS"));
    }

    [Theory]
    [InlineData("bottom:0;width:60px;height:20px", "height:100px", HtmlRenderIntentProfile.ScreenFullPage, 80D, 100D, 200D)]
    [InlineData("", "", HtmlRenderIntentProfile.ScreenFullPage, 80D, 100D, 180D)]
    [InlineData("bottom:0;width:60px;height:20px", "height:100px", HtmlRenderIntentProfile.PrintPaged, 260D, 80D, 80D)]
    [InlineData("", "", HtmlRenderIntentProfile.PrintPaged, 260D, 80D, 60D)]
    public void HtmlColumns_SpannerSourceStateBottomAndStaticAnchorsUseTheTrailingContainingFragment(
        string vertical, string sizing, HtmlRenderIntentProfile profile, double absoluteX, double absoluteY, double afterY) {
        HtmlRenderDocument rendered = RenderDescendantSpanners(
            DescendantSpannerPositionedContent(false, vertical, sizing), profile);

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "ABS", "B1", "B2", "SPAN", "C1", "C2", "T1", "T2", "T3", "T4", "AFTER");
        // Native N33/N34 place bottom-only and automatic vertical insets in
        // the trailing source fragment, independently of top:0's first origin.
        Assert.Equal(absoluteX, DescendantSpannerText(rendered, "ABS").X, 3);
        Assert.Equal(absoluteY, DescendantSpannerText(rendered, "ABS").Y, 3);
        Assert.Equal(afterY, DescendantSpannerText(rendered, "AFTER").Y, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(false, HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData(true, HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData(false, HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    [InlineData(true, HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void HtmlColumns_SpannerSourceStateOrdinaryFiniteOverflowRetainsOpacityPositioningAndFollowingFlow(
        bool nested, HtmlRenderIntentProfile profile) {
        string finite = "<div id='wrapper' style='height:0;padding-top:3px;padding-bottom:5px;"
            + "border-top:2px solid red;border-bottom:2px solid blue;position:relative;opacity:.5'>"
            + "<div style='position:absolute;left:0;top:0;width:80px;height:20px;background:lime'>ABS</div>"
            + "<p>B1</p><p>B2</p><p>S1</p><p>C1</p><p>C2</p></div>";
        HtmlRenderDocument rendered = RenderDescendantSpanners("<p>P1</p><p>P2</p>"
            + (nested ? "<div>" + finite + "</div>" : finite) + DescendantSpannerParagraphs("T", 4),
            profile, extraCss: "#columns{height:100px;background:blue}");

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "ABS", "B1", "B2", "S1", "C1", "C2", "T1", "T2", "T3", "T4", "AFTER");
        // Native N37/N38 continue overflow in column two while later source
        // siblings still advance from the finite 12px box, even through a wrapper.
        Assert.Equal(160D, DescendantSpannerText(rendered, "S1").X, 3);
        Assert.Equal(0D, DescendantSpannerText(rendered, "S1").Y, 3);
        Assert.Equal(20D, DescendantSpannerText(rendered, "C1").Y, 3);
        Assert.Equal(40D, DescendantSpannerText(rendered, "C2").Y, 3);
        Assert.Equal(0D, DescendantSpannerText(rendered, "ABS").X, 3);
        Assert.Equal(42D, DescendantSpannerText(rendered, "ABS").Y, 3);
        Assert.Equal(52D, DescendantSpannerText(rendered, "T1").Y, 3);
        Assert.Equal(100D, DescendantSpannerText(rendered, "AFTER").Y, 3);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0, 128, 127), raster.GetPixel(70, 50));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_SpannerSourceStateStaticSpannerRetainsItsEffectiveAbsoluteContainingBlock() {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<p>P1</p><p>P2</p>"
            + "<div id='wrapper' style='position:relative'><p>B1</p><p>B2</p>"
            + "<div id='spanner' class='spanner'><p>SPAN</p>"
            + "<div id='absolute' style='position:absolute;right:0;top:0;width:20px;height:20px;background:#00ff00'>X</div>"
            + "</div><p>C1</p><p>C2</p></div>" + DescendantSpannerParagraphs("T", 4),
            extraCss: "#columns{position:relative}");

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "B1", "B2", "SPAN", "X", "C1", "C2", "T1", "T2", "T3", "T4", "AFTER");
        HtmlRenderShape absolute = Assert.Single(DescendantSpannerShapes(rendered), shape => shape.Source == "div#absolute"
            && shape.Shape.FillColor == OfficeColor.FromRgb(0, 255, 0));
        // The unpaged CSS witness retains the 300px owner's right edge. The
        // static extracted spanner cannot become a new positioned container.
        Assert.Equal(280D, absolute.X, 3);
        Assert.Equal(0D, absolute.Y, 3);
        Assert.Equal(20D, absolute.Width, 3);
        Assert.Equal(280D, DescendantSpannerText(rendered, "X").X, 3);
        Assert.Equal(300D, DescendantSpannerBackground(rendered, "div#spanner").Width, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_SpannerSourceStateDisplayContentsOpacityDoesNotCompositeItsChildren() {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<div id='wrapper' style='display:contents;opacity:.5;background:red'>"
            + "<p>P1</p><p>P2</p><div id='spanner' class='spanner'><p>SPAN</p></div>"
            + "<p>T1</p><p>T2</p></div>", extraCss: "#columns{background:blue}#spanner{background:red}");

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "SPAN", "T1", "T2", "AFTER");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        // A boxless ancestor contributes neither its red background nor group
        // opacity. The real red spanner therefore paints opaquely over blue.
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(100, 10));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(100, 30));
        Assert.Equal(80D, DescendantSpannerText(rendered, "AFTER").Y, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("height:100px", 120D, 200D)]
    [InlineData("min-height:100px", 120D, 200D)]
    [InlineData("max-height:60px", 80D, 160D)]
    [InlineData("aspect-ratio:7/5", 120D, 200D)]
    public void HtmlColumns_SpannerSourceStateFiniteAncestorHeightIsConsumedOnceAcrossSelections(
        string sizing, double tailY, double afterY) {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<p>P1</p><p>P2</p>"
            + "<div id='wrapper' style='" + sizing + "'><p>B1</p><p>B2</p>"
            + "<h2 id='spanner' class='spanner'>SPAN</h2><p>C1</p><p>C2</p></div>"
            + DescendantSpannerParagraphs("T", 4));

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "B1", "B2", "SPAN", "C1", "C2", "T1", "T2", "T3", "T4", "AFTER");
        // Retained native N22/N26/N27/N28 observations distinguish the target,
        // minimum, maximum and aspect-ratio allocation of the original box.
        HtmlRenderShape spanner = DescendantSpannerBackground(rendered, "h2#spanner");
        Assert.Equal(300D, spanner.Width, 3);
        Assert.Equal(40D, spanner.Y, 3);
        Assert.Equal(60D, DescendantSpannerText(rendered, "C1").Y, 3);
        Assert.Equal(80D, DescendantSpannerText(rendered, "C2").Y, 3);
        Assert.Equal(tailY, DescendantSpannerText(rendered, "T1").Y, 3);
        Assert.Equal(afterY, DescendantSpannerText(rendered, "AFTER").Y, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_SpannerSourceStateFiniteOverflowPaintContinuesWithoutAdvancingFollowingFlow(
        HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<p>P1</p><p>P2</p>"
            + "<div id='wrapper' style='height:0;padding-top:3px;padding-bottom:5px;"
            + "border-top:2px solid red;border-bottom:2px solid blue'><p>B1</p><p>B2</p>"
            + "<h2 id='spanner' class='spanner'>SPAN</h2><p>C1</p><p>C2</p></div>"
            + DescendantSpannerParagraphs("T", 4), profile);

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "B1", "B2", "SPAN", "C1", "C2", "T1", "T2", "T3", "T4", "AFTER");
        // Native N29 keeps the 7px closing box advance while its second line
        // continues in column two. That overflow must not consume T1's space.
        Assert.Equal(65D, DescendantSpannerText(rendered, "SPAN").Y, 3);
        Assert.Equal(20D, DescendantSpannerText(rendered, "C1").X, 3);
        Assert.Equal(85D, DescendantSpannerText(rendered, "C1").Y, 3);
        Assert.Equal(180D, DescendantSpannerText(rendered, "C2").X, 3);
        Assert.Equal(85D, DescendantSpannerText(rendered, "C2").Y, 3);
        Assert.Equal(1, DescendantSpannerTextPage(rendered, "C2").PageNumber);
        Assert.Equal(92D, DescendantSpannerText(rendered, "T1").Y, 3);
        Assert.Equal(85D, DescendantSpannerText(rendered, "T2").Y, 3);
        Assert.Equal(2, DescendantSpannerTextPage(rendered, "T3").PageNumber);
        Assert.Equal(60D, DescendantSpannerText(rendered, "AFTER").Y, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_SpannerSourceStateClippedAncestorRetainsLogicalSourceWithoutLeakingSpannerPaint() {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<p>P1</p><p>P2</p>"
            + "<div id='wrapper' style='clip-path:inset(0 100% 0 0)'><p>B1</p><p>B2</p>"
            + "<h2 id='spanner' class='spanner'>SPAN</h2><p>C1</p><p>C2</p></div>"
            + DescendantSpannerParagraphs("T", 4), extraCss: "#columns{background:blue}#spanner{background:red}");

        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "B1", "B2", "SPAN", "C1", "C2", "T1", "T2", "T3", "T4", "AFTER");
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported
            && diagnostic.Source == "div#wrapper" && diagnostic.Detail?.Contains("clip-path", StringComparison.Ordinal) == true);
        Assert.True(rendered.HasLoss);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        Assert.Contains(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderPathClipGroup>(),
            clip => clip.ClipPath.Kind == OfficeClipPathKind.Empty
                && DescendantSpannerOwnedText(clip.Visuals).Any(text => text.Text == "SPAN"));
        HtmlRenderShape spanner = Assert.Single(DescendantSpannerShapes(rendered), shape => shape.Source == "h2#spanner"
            && shape.Shape.FillColor == OfficeColor.Red);
        Assert.Equal(140D, spanner.Width, 3);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        int x = (int)Math.Floor(spanner.X + spanner.Width - 10D), y = (int)Math.Floor(spanner.Y + 5D);
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(x, y));
    }

    private static string DescendantSpannerPositionedContent(bool beforeSpanner,
        string vertical = "top:0", string sizing = "") {
        string positioned = "<div id='positioned' style='position:absolute;left:80px;" + vertical + "'>ABS</div>";
        return "<p>P1</p><p>P2</p><div id='wrapper' style='position:relative;" + sizing + "'>"
            + (beforeSpanner ? positioned : string.Empty)
            + "<p>B1</p><p>B2</p><h2 id='spanner' class='spanner'>SPAN</h2><p>C1</p><p>C2</p>"
            + (beforeSpanner ? string.Empty : positioned) + "</div>" + DescendantSpannerParagraphs("T", 4);
    }
}
