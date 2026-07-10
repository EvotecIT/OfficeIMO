using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlFlexRow_AppliesGapMainDistributionAndCrossAlignment() {
        const string html = """
            <div id="flex" style="display:flex;width:300px;height:80px;gap:10px;justify-content:space-between;align-items:center">
              <div id="a" style="width:50px;height:20px;background:#ff0000">A</div>
              <div id="b" style="width:70px;height:40px;background:#0000ff">B</div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 400D);
        HtmlRenderShape first = FindFlexShape(rendered, "div#a");
        HtmlRenderShape second = FindFlexShape(rendered, "div#b");

        Assert.Equal(0D, first.X, 3);
        Assert.Equal(30D, first.Y, 3);
        Assert.Equal(50D, first.Width, 3);
        Assert.Equal(20D, first.Height, 3);
        Assert.Equal(230D, second.X, 3);
        Assert.Equal(20D, second.Y, 3);
        Assert.Equal(70D, second.Width, 3);
        Assert.Equal(40D, second.Height, 3);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexRow_DistributesGrowAndShrinkFromTheFlexBasis() {
        const string html = """
            <div style="display:flex;width:300px">
              <div id="grow-one" style="flex:1 1 0%;height:20px;background:#ff0000"></div>
              <div id="grow-two" style="flex:2 1 0%;height:20px;background:#0000ff"></div>
            </div>
            <div style="display:flex;width:300px">
              <div id="shrink-one" style="flex:0 1 200px;height:20px;background:#00ff00"></div>
              <div id="shrink-two" style="flex:0 1 200px;height:20px;background:#ffff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 320D);

        Assert.Equal(100D, FindFlexShape(rendered, "div#grow-one").Width, 3);
        Assert.Equal(200D, FindFlexShape(rendered, "div#grow-two").Width, 3);
        Assert.Equal(150D, FindFlexShape(rendered, "div#shrink-one").Width, 3);
        Assert.Equal(150D, FindFlexShape(rendered, "div#shrink-two").Width, 3);
    }

    [Fact]
    public void HtmlFlexRow_RespectsMinAndMaxConstraintsDuringDistribution() {
        const string html = """
            <div style="display:flex;width:300px">
              <div id="min" style="flex:0 1 200px;min-width:180px;height:20px;background:#ff0000"></div>
              <div id="after-min" style="flex:0 1 200px;height:20px;background:#0000ff"></div>
            </div>
            <div style="display:flex;width:300px">
              <div id="max" style="flex:1 1 0%;max-width:80px;height:20px;background:#00ff00"></div>
              <div id="after-max" style="flex:1 1 0%;height:20px;background:#ffff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 320D);

        Assert.Equal(180D, FindFlexShape(rendered, "div#min").Width, 3);
        Assert.Equal(120D, FindFlexShape(rendered, "div#after-min").Width, 3);
        Assert.Equal(80D, FindFlexShape(rendered, "div#max").Width, 3);
        Assert.Equal(220D, FindFlexShape(rendered, "div#after-max").Width, 3);
    }

    [Fact]
    public void HtmlFlexRow_CombinesOrderWithRowReverseWithoutChangingPaintOrder() {
        const string html = """
            <div style="display:flex;flex-direction:row-reverse;width:200px">
              <div id="a" style="order:2;width:50px;height:20px;background:#ff0000"></div>
              <div id="b" style="order:1;width:50px;height:20px;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 220D);
        HtmlRenderShape firstInPaintOrder = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().First(shape => shape.Source == "div#b" || shape.Source == "div#a");
        HtmlRenderShape a = FindFlexShape(rendered, "div#a");
        HtmlRenderShape b = FindFlexShape(rendered, "div#b");

        Assert.Equal("div#b", firstInPaintOrder.Source);
        Assert.Equal(100D, a.X, 3);
        Assert.Equal(150D, b.X, 3);
    }

    [Fact]
    public void HtmlFlexRow_StretchesAutoCrossSizesAndHonorsAlignSelf() {
        const string html = """
            <div style="display:flex;width:200px;height:100px;align-items:stretch">
              <div id="stretch" style="width:60px;background:#ff0000">A</div>
              <div id="end" style="width:60px;height:20px;align-self:flex-end;background:#0000ff">B</div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 220D);
        HtmlRenderShape stretched = FindFlexShape(rendered, "div#stretch");
        HtmlRenderShape alignedEnd = FindFlexShape(rendered, "div#end");

        Assert.Equal(100D, stretched.Height, 3);
        Assert.Equal(0D, stretched.Y, 3);
        Assert.Equal(20D, alignedEnd.Height, 3);
        Assert.Equal(80D, alignedEnd.Y, 3);
    }

    [Fact]
    public void HtmlFlexRow_ComposesNestedFlexContainersWithinAllocatedItems() {
        const string html = """
            <div id="outer" style="display:flex;width:240px;height:40px">
              <div id="left" style="flex:1 1 0%;height:40px;background:#eeeeee">
                <div id="inner" style="display:flex;width:100%;height:40px;justify-content:space-between">
                  <div id="inner-a" style="width:40px;height:40px;background:#ff0000"></div>
                  <div id="inner-b" style="width:40px;height:40px;background:#0000ff"></div>
                </div>
              </div>
              <div id="right" style="flex:1 1 0%;height:40px;background:#00ff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 260D);

        Assert.Equal(0D, FindFlexShape(rendered, "div#inner-a").X, 3);
        Assert.Equal(80D, FindFlexShape(rendered, "div#inner-b").X, 3);
        Assert.Equal(120D, FindFlexShape(rendered, "div#right").X, 3);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexRow_MovesAsOneUnitWhenItFitsOnlyOnTheNextPage() {
        const string html = """
            <div style="height:60px;margin:0">Before</div>
            <div id="flex" style="display:flex;width:160px;height:50px">
              <div id="page-a" style="width:80px;height:50px;background:#ff0000">A</div>
              <div id="page-b" style="width:80px;height:50px;background:#0000ff">B</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#page-a" || shape.Source == "div#page-b");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#page-a");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#page-b");
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlFlexRow_DiagnosesUnsupportedValuesWithoutDiscardingItems() {
        const string html = """
            <div id="flex" style="display:flex;width:200px;gap:calc(4px + 2px);justify-content:safe center">
              <div id="item" style="flex-basis:calc(20px + 5px);height:20px;background:#ff0000">Item</div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 220D);

        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Item");
        Assert.Equal(3, rendered.Diagnostics.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexValueUnsupported));
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexRow_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = """
            <div style="display:flex;width:60px;height:20px;gap:10px">
              <div style="width:20px;height:20px;background:#ff0000"></div>
              <div style="width:20px;height:20px;background:#0000ff"></div>
            </div>
            <p style="margin:0">FlexPdfMarker</p>
            """;
        var options = new HtmlImageExportOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = html.ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(html.SaveAsPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(10, 10));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(40, 10));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"8\" y=\"8\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("<rect x=\"38\" y=\"8\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("FlexPdfMarker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlFlexFallbacks_RemainExplicitForWrapDirectTextAutoMarginsAndInlineFlex() {
        const string html = """
            <div class="wrap" style="display:flex;flex-wrap:wrap"><span>Wrap</span></div>
            <div class="text" style="display:flex">Direct text</div>
            <div class="margin" style="display:flex"><span style="margin-left:auto">Auto</span></div>
            <span class="inline" style="display:inline-flex"><span>Inline</span></span>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 240D);
        IReadOnlyList<HtmlDiagnostic> diagnostics = rendered.Diagnostics.Diagnostics
            .Where(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending)
            .ToList();

        Assert.Equal(4, diagnostics.Count);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "div.wrap");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "div.text");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "div.margin");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "span.inline");
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Direct", StringComparison.Ordinal));
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Auto");
    }

    private static HtmlRenderDocument RenderFlex(string html, double viewportWidth) =>
        HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = viewportWidth,
            Margins = HtmlRenderMargins.All(0D)
        });

    private static HtmlRenderShape FindFlexShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);
}
