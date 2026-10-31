using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlFixedHeightOverflow_PagesVisibleContentWithoutAdvancingTheSignedFollowingFlow(
        HtmlRenderIntentProfile profile) {
        const string link = "https://example.test/overflow-last";
        string html = FixedHeightOverflowSource("<div style='height:1060px'>FIXED-LEAD</div>"
            + "<div id='short' style='height:25px;margin-bottom:-45px;background:#00aa00'>"
            + "<p>FIXED-X01</p><p>FIXED-X02</p><p><a href='" + link + "'>FIXED-X03</a></p></div>"
            + "<p>FIXED-AFTER</p>");
        HtmlPdfRenderRequestResult output = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, FixedHeightOverflowOptions()));
        HtmlRenderDocument rendered = output.RenderResult.Document;

        Assert.Equal(12, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "FIXED-LEAD", "FIXED-X01", "FIXED-X02", "FIXED-X03", "FIXED-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "FIXED-LEAD", 1, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "FIXED-X01", 11, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "FIXED-X02", 11, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "FIXED-X03", 12, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "FIXED-AFTER", 11, 40D);
        Assert.Equal(link, DescendantSpannerText(rendered, "FIXED-X03").LinkUri);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);

        OfficeRasterImage penultimate = OfficeDrawingRasterRenderer.Render(rendered.Pages[10].CreateDrawing());
        OfficeRasterImage last = OfficeDrawingRasterRenderer.Render(rendered.Pages[11].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0, 170, 0), penultimate.GetPixel(500, 70));
        Assert.Equal(OfficeColor.White, last.GetPixel(500, 10));

        PdfCore.PdfReadDocument pdf = PdfCore.PdfReadDocument.Open(output.ToBytes());
        Assert.Equal(12, pdf.Pages.Count);
        string[] pdfWords = pdf.ExtractText().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        // The following paragraph is physically on page 11, before the final
        // overflow line on page 12. Complete source counts do not imply page order.
        Assert.Equal(new[] { "FIXED-LEAD", "FIXED-X01", "FIXED-X02", "FIXED-X03", "FIXED-AFTER" }
            .OrderBy(word => word, StringComparer.Ordinal), pdfWords.OrderBy(word => word, StringComparer.Ordinal));
        Assert.Single(pdf.Pages.SelectMany(page => page.GetLinkAnnotations()), annotation => annotation.Uri == link);
    }

    [Theory]
    [InlineData("height:20px")]
    [InlineData("max-height:20px")]
    [InlineData("aspect-ratio:8;min-height:0")]
    public void HtmlFixedHeightOverflow_ConstrainedBlockRetainsItsOverflowWithoutGrowingFollowingFlow(string constraint) {
        string html = FixedHeightOverflowSource("<div style='height:60px'>FLOW-LEAD</div>"
            + "<div style='width:160px;" + constraint + "'><p>FLOW-X01</p><p>FLOW-X02</p><p>FLOW-X03</p></div>"
            + "<p style='padding-left:300px'>FLOW-AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "FLOW-LEAD", "FLOW-X01", "FLOW-X02", "FLOW-X03", "FLOW-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "FLOW-X01", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "FLOW-X02", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "FLOW-X03", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "FLOW-AFTER", 1, 80D);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Theory]
    [InlineData("nested")]
    [InlineData("inline-sibling")]
    public void HtmlFixedHeightOverflow_PreservesPaintAcrossOrdinaryWrappersAndInlineSiblingFlow(string route) {
        string content = "<div style='height:60px'>WRAP-LEAD</div>"
            + "<div style='height:20px;margin-bottom:-40px'><p>WRAP-X01</p>"
            + "<p>WRAP-X02</p><p>WRAP-X03</p></div>"
            + (route == "inline-sibling" ? "<span>WRAP-AFTER</span>" : "<p>WRAP-AFTER</p>");
        if (route == "nested") content = "<section>" + content + "</section>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            FixedHeightOverflowSource(content), FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "WRAP-LEAD", "WRAP-X01", "WRAP-X02", "WRAP-X03", "WRAP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "WRAP-X01", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "WRAP-X02", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "WRAP-X03", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "WRAP-AFTER", 1, 40D);
    }

    [Theory]
    [InlineData("hidden")]
    [InlineData("clip")]
    public void HtmlFixedHeightOverflow_AuthoredClipDoesNotAllocatePagesForTheHiddenTail(string overflow) {
        string html = FixedHeightOverflowSource("<div style='height:60px'>CLIP-LEAD</div>"
            + "<div style='height:20px;margin-bottom:-40px;overflow:" + overflow + "'>"
            + "<p style='background:#00aa00'>CLIP-X01</p><p style='background:#00aa00'>CLIP-X02</p>"
            + "<p style='background:#00aa00'>CLIP-X03</p></div><p>CLIP-AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        HtmlRenderPage page = Assert.Single(rendered.Pages);
        AssertFixedHeightOverflowPlacement(rendered, "CLIP-X01", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "CLIP-AFTER", 1, 40D);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0, 170, 0), image.GetPixel(500, 70));
        Assert.Equal(OfficeColor.White, image.GetPixel(500, 90));
    }

    [Theory]
    [InlineData("break-inside:avoid;orphans:1;widows:1")]
    [InlineData("orphans:2;widows:2")]
    public void HtmlFixedHeightOverflow_ParagraphKeepInTheOverflowDoesNotDeferFollowingFlow(string keep) {
        string html = FixedHeightOverflowSource("<div style='height:60px'>KEEP-LEAD</div>"
            + "<div style='height:20px'><p>KEEP-INNER</p><p style='" + keep + "'>"
            + "KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 40D);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Theory]
    [InlineData("min-height:0")]
    [InlineData("max-height:9999px")]
    public void HtmlFixedHeightOverflow_NonbindingWrapperConstraintRetainsIndependentOverflowFlow(string constraint) {
        string html = FixedHeightOverflowSource("<section style='" + constraint + "'>"
            + "<div style='height:60px'>KEEP-LEAD</div><div style='height:20px'><p>KEEP-INNER</p>"
            + "<p style='break-inside:avoid;orphans:1;widows:1'>KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 40D);
    }

    [Theory]
    [InlineData(80, 40D)]
    [InlineData(100, 60D)]
    public void HtmlFixedHeightOverflow_InteriorKeepGapCountsTowardTheAuthoredHeight(int height, double followingTop) {
        string html = FixedHeightOverflowSource("<div style='height:60px'>KEEP-LEAD</div>"
            + "<div style='height:" + height + "px'><p>KEEP-INNER</p>"
            + "<p style='break-inside:avoid;orphans:1;widows:1'>KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 40D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 2, followingTop);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Theory]
    [InlineData(40, 20, 20)]
    [InlineData(80, -20, -20)]
    public void HtmlFixedHeightOverflow_CollapsedMarginsSetTheKeepProjectionOrigin(
        int leadHeight, int precedingMargin, int followingMargin) {
        string html = FixedHeightOverflowSource("<div style='height:" + leadHeight + "px;margin-bottom:"
            + precedingMargin + "px'>KEEP-LEAD</div><div style='height:20px;margin-top:"
            + followingMargin + "px'><p>KEEP-INNER</p>"
            + "<p style='break-inside:avoid;orphans:1;widows:1'>KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 40D);
    }

    [Fact]
    public void HtmlFixedHeightOverflow_ParagraphBackgroundFillsTheGapBeforeItsNextFragment() {
        string html = FixedHeightOverflowSource("<div style='height:70px'>PAINT-LEAD</div>"
            + "<div style='height:20px'><p style='background:#00aa00'>PAINT-1<br>PAINT-2<br>PAINT-3</p></div>"
            + "<p style='padding-left:300px'>PAINT-AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "PAINT-LEAD", "PAINT-1", "PAINT-2", "PAINT-3", "PAINT-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "PAINT-1", 1, 70D);
        AssertFixedHeightOverflowPlacement(rendered, "PAINT-2", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "PAINT-3", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "PAINT-AFTER", 2, 0D);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0, 170, 0), first.GetPixel(500, 95));
    }

    [Fact]
    public void HtmlFixedHeightOverflow_ParallelHiddenKeepRetainsItsOwnVisibleFragment() {
        string html = FixedHeightOverflowSource("<div style='height:60px'>LEAD</div>"
            + "<div style='height:20px'><p>X1</p><p>X2</p><p>X3</p></div>"
            + "<div style='height:20px;overflow:hidden'><p style='background:#00aa00;break-inside:avoid;'>"
            + "CLIP1<br>CLIP2<br>CLIP3</p></div><p>AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertFixedHeightOverflowPlacement(rendered, "X1", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "X2", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "X3", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "CLIP1", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "AFTER", 2, 0D);
        string[] logical = rendered.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        foreach (string label in new[] { "LEAD", "X1", "X2", "X3", "CLIP1", "AFTER" })
            Assert.Single(logical, word => word == label);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeRasterImage second = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0, 170, 0), first.GetPixel(500, 90));
        Assert.Equal(OfficeColor.White, second.GetPixel(500, 10));
    }

    private static string FixedHeightOverflowSource(string content) => "<!doctype html><style>"
        + "html,body,p{margin:0;padding:0}html,body{font:14px/20px Arial;background:white}"
        + "p{font:inherit;orphans:1;widows:1}</style>" + content;

    private static HtmlRenderOptions FixedHeightOverflowOptions() {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(700D / HtmlRenderOptions.CssPixelsPerInch), global::OfficeIMO.OfficeLength.Inches(100D / HtmlRenderOptions.CssPixelsPerInch)),
            ViewportWidth = 800D,
            ViewportHeight = 1000D,
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            AllowSystemFontFallback = false,
            BackgroundColor = OfficeColor.White
        };
        options.Fonts.Add("Arial", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            Enumerable.Range(32, 95).ToArray()));
        return options;
    }

    private static void AssertFixedHeightOverflowPlacement(HtmlRenderDocument rendered, string label, int page, double top) {
        Assert.Equal(page, DescendantSpannerTextPage(rendered, label).PageNumber);
        Assert.Equal(top, DescendantSpannerText(rendered, label).Y, 3);
    }
}
