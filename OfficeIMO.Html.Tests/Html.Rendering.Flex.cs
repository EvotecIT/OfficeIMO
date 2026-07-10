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
    public void HtmlFlexWrap_CreatesLinesWithIndependentMainAndCrossAlignment() {
        const string html = """
            <div style="display:flex;flex-wrap:wrap;width:120px;gap:5px 10px;justify-content:space-between;align-items:center">
              <div id="wrap-a" style="width:50px;height:20px;background:#ff0000"></div>
              <div id="wrap-b" style="width:50px;height:30px;background:#0000ff"></div>
              <div id="wrap-c" style="width:50px;height:10px;background:#00ff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 140D);
        HtmlRenderShape a = FindFlexShape(rendered, "div#wrap-a");
        HtmlRenderShape b = FindFlexShape(rendered, "div#wrap-b");
        HtmlRenderShape c = FindFlexShape(rendered, "div#wrap-c");

        Assert.Equal(0D, a.X, 3);
        Assert.Equal(5D, a.Y, 3);
        Assert.Equal(70D, b.X, 3);
        Assert.Equal(0D, b.Y, 3);
        Assert.Equal(0D, c.X, 3);
        Assert.Equal(35D, c.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexWrapReverse_UsesTheReversedCrossStartWithAlignContent() {
        const string html = """
            <div style="display:flex;flex-wrap:wrap-reverse;width:100px;height:100px;row-gap:10px;align-content:center">
              <div id="reverse-a" style="width:60px;height:20px;background:#ff0000"></div>
              <div id="reverse-b" style="width:60px;height:30px;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 120D);

        Assert.Equal(60D, FindFlexShape(rendered, "div#reverse-a").Y, 3);
        Assert.Equal(20D, FindFlexShape(rendered, "div#reverse-b").Y, 3);
    }

    [Fact]
    public void HtmlFlexWrap_DiagnosesUnsupportedRowGapAndReverseOverflow() {
        const string html = """
            <div style="display:flex;flex-wrap:wrap-reverse;width:100px;height:30px;row-gap:calc(2px + 1px)">
              <div id="overflow-a" style="width:100px;height:20px;background:#ff0000"></div>
              <div id="overflow-b" style="width:100px;height:20px;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 120D);
        IReadOnlyList<HtmlDiagnostic> diagnostics = rendered.Diagnostics.Diagnostics
            .Where(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexValueUnsupported)
            .ToList();

        Assert.Equal(2, diagnostics.Count);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Detail == "row-gap=calc(2px + 1px)");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Detail == "flex-wrap=wrap-reverse; cross-size-overflow");
        Assert.True(FindFlexShape(rendered, "div#overflow-a").Y >= 0D);
        Assert.True(FindFlexShape(rendered, "div#overflow-b").Y >= 0D);
    }

    [Fact]
    public void HtmlFlexWrap_PaginatesOnlyBetweenCompleteLines() {
        const string html = """
            <div style="height:20px;margin:0">Before</div>
            <div style="display:flex;flex-wrap:wrap;width:100px">
              <div id="line-one" style="width:100px;height:40px;background:#ff0000">One</div>
              <div id="line-two" style="width:100px;height:40px;background:#0000ff">Two</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 70D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#line-one");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#line-two");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#line-two");
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlFlexColumn_AppliesMainDistributionAndCrossAlignment() {
        const string html = """
            <div style="display:flex;flex-direction:column;width:100px;height:300px;gap:10px;justify-content:space-between;align-items:center">
              <div id="column-a" style="width:20px;height:50px;background:#ff0000"></div>
              <div id="column-b" style="width:40px;height:70px;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 120D);
        HtmlRenderShape a = FindFlexShape(rendered, "div#column-a");
        HtmlRenderShape b = FindFlexShape(rendered, "div#column-b");

        Assert.Equal(40D, a.X, 3);
        Assert.Equal(0D, a.Y, 3);
        Assert.Equal(20D, a.Width, 3);
        Assert.Equal(50D, a.Height, 3);
        Assert.Equal(30D, b.X, 3);
        Assert.Equal(230D, b.Y, 3);
        Assert.Equal(40D, b.Width, 3);
        Assert.Equal(70D, b.Height, 3);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexColumn_DistributesGrowShrinkAndVerticalConstraints() {
        const string html = """
            <div style="display:flex;flex-direction:column;width:80px;height:300px">
              <div id="column-grow-one" style="flex:1 1 0%;max-height:80px;background:#ff0000"></div>
              <div id="column-grow-two" style="flex:1 1 0%;background:#0000ff"></div>
            </div>
            <div style="display:flex;flex-direction:column;width:80px;height:300px">
              <div id="column-shrink-one" style="flex:0 1 200px;min-height:180px;background:#00ff00"></div>
              <div id="column-shrink-two" style="flex:0 1 200px;background:#ffff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 100D);

        Assert.Equal(80D, FindFlexShape(rendered, "div#column-grow-one").Height, 3);
        Assert.Equal(220D, FindFlexShape(rendered, "div#column-grow-two").Height, 3);
        Assert.Equal(180D, FindFlexShape(rendered, "div#column-shrink-one").Height, 3);
        Assert.Equal(120D, FindFlexShape(rendered, "div#column-shrink-two").Height, 3);
    }

    [Fact]
    public void HtmlFlexColumn_ResolvesPercentageBasisOnlyAgainstDefiniteHeight() {
        const string html = """
            <div style="display:flex;flex-direction:column;width:80px;height:200px">
              <div id="definite-quarter" style="flex:0 0 25%;background:#ff0000"></div>
              <div id="definite-rest" style="flex:0 0 75%;background:#0000ff"></div>
            </div>
            <div style="display:flex;flex-direction:column;width:80px">
              <div id="indefinite-a" style="flex:0 0 50%;height:30px;background:#00ff00"></div>
              <div id="indefinite-b" style="flex:0 0 50%;height:40px;background:#ffff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 100D);

        Assert.Equal(50D, FindFlexShape(rendered, "div#definite-quarter").Height, 3);
        Assert.Equal(150D, FindFlexShape(rendered, "div#definite-rest").Height, 3);
        Assert.Equal(30D, FindFlexShape(rendered, "div#indefinite-a").Height, 3);
        Assert.Equal(40D, FindFlexShape(rendered, "div#indefinite-b").Height, 3);
    }

    [Fact]
    public void HtmlFlexColumnReverse_CombinesOrderWithReversedMainPlacement() {
        const string html = """
            <div style="display:flex;flex-direction:column-reverse;width:80px;height:200px">
              <div id="column-reverse-a" style="order:2;height:50px;background:#ff0000"></div>
              <div id="column-reverse-b" style="order:1;height:50px;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 100D);
        HtmlRenderShape firstInPaintOrder = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .First(shape => shape.Source == "div#column-reverse-a" || shape.Source == "div#column-reverse-b");

        Assert.Equal("div#column-reverse-b", firstInPaintOrder.Source);
        Assert.Equal(100D, FindFlexShape(rendered, "div#column-reverse-a").Y, 3);
        Assert.Equal(150D, FindFlexShape(rendered, "div#column-reverse-b").Y, 3);
    }

    [Fact]
    public void HtmlFlexColumn_PaginatesOnlyBetweenCompleteItems() {
        const string html = """
            <div style="height:20px;margin:0">Before</div>
            <div style="display:flex;flex-direction:column;width:100px">
              <div id="column-page-one" style="height:40px;background:#ff0000">One</div>
              <div id="column-page-two" style="height:40px;background:#0000ff">Two</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 70D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#column-page-one");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#column-page-two");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#column-page-two");
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlFlexColumn_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = """
            <div style="display:flex;flex-direction:column;width:20px;height:50px;gap:10px">
              <div style="width:20px;height:20px;background:#ff0000"></div>
              <div style="width:20px;height:20px;background:#0000ff"></div>
            </div>
            <p style="margin:0">ColumnFlexPdfMarker</p>
            """;
        var options = new HtmlImageExportOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = html.ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(html.SaveAsPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(10, 10));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(10, 40));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"8\" y=\"8\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("<rect x=\"8\" y=\"38\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("ColumnFlexPdfMarker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
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
    public void HtmlFlexFallbacks_RemainExplicitForColumnWrapDirectTextAutoMarginsAndInlineFlex() {
        const string html = """
            <div class="column-wrap" style="display:flex;flex-direction:column;flex-wrap:wrap"><span>Column wrap</span></div>
            <div class="text" style="display:flex">Direct text</div>
            <div class="margin" style="display:flex"><span style="margin-left:auto">Auto</span></div>
            <span class="inline" style="display:inline-flex"><span>Inline</span></span>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 240D);
        IReadOnlyList<HtmlDiagnostic> diagnostics = rendered.Diagnostics.Diagnostics
            .Where(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending)
            .ToList();

        Assert.Equal(4, diagnostics.Count);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "div.column-wrap");
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
