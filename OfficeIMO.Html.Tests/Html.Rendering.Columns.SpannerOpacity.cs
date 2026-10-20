using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlColumns_DescendantOpacityIncludesTheExtractedSpannerWithoutDoubleCompositing() {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<div id='wrapper'><p>P1</p><p>P2</p>"
            + "<h2 id='spanner' class='spanner'>SPAN</h2><p>T1</p><p>T2</p></div>",
            extraCss: "#columns{background:blue}#wrapper{background:red;opacity:.5}#spanner{background:red}");
        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "SPAN", "T1", "T2", "AFTER");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        // CSS group opacity includes the spanner. Sampling away from ink also
        // catches an opaque extracted box or opacity applied to each layer twice.
        OfficeColor ordinary = raster.GetPixel(100, 10), spanner = raster.GetPixel(100, 30);
        Assert.Equal(ordinary, spanner);
        Assert.InRange(spanner.R, (byte)126, (byte)129);
        Assert.Equal((byte)0, spanner.G);
        Assert.InRange(spanner.B, (byte)126, (byte)129);
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_SpannerSourceFragmentsRetainInheritedLinksAndOneAncestorDestination() {
        HtmlRenderDocument rendered = RenderDescendantSpanners("<a id='wrapper' href='#after' style='display:block'>"
            + "<p>P1</p><p>P2</p><h2 id='spanner' class='spanner'>SPAN</h2><p>T1</p><p>T2</p></a>",
            entry: "<p><a href='#wrapper'>JUMP</a></p>");
        AssertDescendantSpannerSourceOnce(rendered, "JUMP", "P1", "P2", "SPAN", "T1", "T2", "AFTER");
        Assert.Equal("#after", DescendantSpannerText(rendered, "SPAN").LinkUri);
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).ToArray();
        Assert.Single(visuals.OfType<HtmlRenderNamedDestination>(), destination => destination.Name == "wrapper");
        Assert.Single(visuals.OfType<HtmlRenderNamedDestination>(), destination => destination.Name == "after");
        string?[] keys = visuals.OfType<HtmlRenderSemanticGroup>().Where(group => group.Source == "a#wrapper")
            .Select(group => group.StructureElementKey).Distinct(StringComparer.Ordinal).ToArray();
        Assert.Single(keys);
        Assert.False(string.IsNullOrEmpty(keys[0]));
        AssertColumnBreaksQualified(rendered);
    }
}
