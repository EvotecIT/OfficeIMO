using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlFlexColumn_NestedDefaultLayoutsRemainLinear() {
        var html = new StringBuilder();
        for (int index = 0; index < 24; index++) html.Append("<div style='display:flex;flex-direction:column'>");
        html.Append("<span>LinearLeaf</span>");
        for (int index = 0; index < 24; index++) html.Append("</div>");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html.ToString(), new HtmlRenderOptions {
            ViewportWidth = 200D,
            MaxLayoutOperations = 100
        });

        Assert.Contains("LinearLeaf", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlFlexColumn_StopsRepeatedReflowAtOperationLimit() {
        var html = new StringBuilder();
        for (int index = 0; index < 12; index++) html.Append("<div style='display:flex;flex-direction:column;align-items:flex-start'>");
        html.Append("<span>BoundedLeaf</span>");
        for (int index = 0; index < 12; index++) html.Append("</div>");

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(html.ToString(), new HtmlRenderOptions { MaxLayoutOperations = 20 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
    }

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
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexRow_ResolvesPercentageHeightsAgainstADefiniteParentHeight() {
        HtmlRenderDocument rendered = RenderFlex("""
            <div id="chart" style="display:flex;align-items:flex-end;width:300px;height:110px">
              <div id="bar" style="width:40px;height:42%;background:#2563eb"></div>
            </div>
            """, 320D);

        HtmlRenderShape bar = FindFlexShape(rendered, "div#bar");

        Assert.Equal(46.2D, bar.Height, 3);
        Assert.Equal(63.8D, bar.Y, 3);
    }

    [Fact]
    public void HtmlPercentageHeight_RemainsContentDrivenWhenTheParentHeightIsIndefinite() {
        HtmlRenderDocument rendered = RenderFlex("""
            <div style="width:300px">
              <div id="content-height" style="height:50%;background:#2563eb">Marker</div>
            </div>
            """, 320D);

        HtmlRenderShape child = FindFlexShape(rendered, "div#content-height");
        Assert.True(child.Height < 40D);
        Assert.True(child.Height > 10D);
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
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#page-a" || shape.Source == "div#page-b");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#page-a");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#page-b");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
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
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
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
        IReadOnlyList<HtmlDiagnostic> diagnostics = rendered.Diagnostics
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#line-one");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#line-two");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#line-two");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
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
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexColumn_PaginatesInsideOneOversizedItemAtNestedBlockBoundaries() {
        const string html = """
            <div id="column" style="display:flex;flex-direction:column;width:100px">
              <div id="oversized" style="background:#eeeeee">
                <div style="height:20px">One</div><div style="height:20px">Two</div>
                <div style="height:20px">Three</div><div style="height:20px">Four</div>
                <div style="height:20px">Five</div><div style="height:20px">Six</div>
              </div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 50D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => Assert.True(page.Visuals.Count > 1));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#column-page-one");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#column-page-two");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#column-page-two");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
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
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(10, 10));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(10, 40));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"8\" y=\"8\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("<rect x=\"8\" y=\"38\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("ColumnFlexPdfMarker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlFlexColumnWrap_CreatesColumnsAndDistributesTheirCrossAxis() {
        const string html = """
            <div style="display:flex;flex-direction:column;flex-wrap:wrap;width:220px;height:120px;gap:10px 20px;align-content:space-between;align-items:flex-start">
              <div id="column-wrap-a" style="width:50px;height:50px;background:#ff0000"></div>
              <div id="column-wrap-b" style="width:50px;height:50px;background:#0000ff"></div>
              <div id="column-wrap-c" style="width:50px;height:50px;background:#00ff00"></div>
              <div id="column-wrap-d" style="width:50px;height:50px;background:#ffff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 240D);

        Assert.Equal(0D, FindFlexShape(rendered, "div#column-wrap-a").X, 3);
        Assert.Equal(0D, FindFlexShape(rendered, "div#column-wrap-a").Y, 3);
        Assert.Equal(0D, FindFlexShape(rendered, "div#column-wrap-b").X, 3);
        Assert.Equal(60D, FindFlexShape(rendered, "div#column-wrap-b").Y, 3);
        Assert.Equal(170D, FindFlexShape(rendered, "div#column-wrap-c").X, 3);
        Assert.Equal(0D, FindFlexShape(rendered, "div#column-wrap-c").Y, 3);
        Assert.Equal(170D, FindFlexShape(rendered, "div#column-wrap-d").X, 3);
        Assert.Equal(60D, FindFlexShape(rendered, "div#column-wrap-d").Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexColumnWrapReverse_ReversesColumnsAndGrowsItemsPerColumn() {
        const string html = """
            <div style="display:flex;flex-direction:column;flex-wrap:wrap-reverse;width:220px;height:120px;gap:10px 20px;align-content:space-between;align-items:flex-start">
              <div id="column-reverse-wrap-a" style="flex:1 1 40px;width:50px;background:#ff0000"></div>
              <div id="column-reverse-wrap-b" style="flex:1 1 40px;width:50px;background:#0000ff"></div>
              <div id="column-reverse-wrap-c" style="flex:1 1 40px;width:50px;background:#00ff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 240D);
        HtmlRenderShape a = FindFlexShape(rendered, "div#column-reverse-wrap-a");
        HtmlRenderShape b = FindFlexShape(rendered, "div#column-reverse-wrap-b");
        HtmlRenderShape c = FindFlexShape(rendered, "div#column-reverse-wrap-c");

        Assert.Equal(170D, a.X, 3);
        Assert.Equal(55D, a.Height, 3);
        Assert.Equal(65D, b.Y, 3);
        Assert.Equal(55D, b.Height, 3);
        Assert.Equal(0D, c.X, 3);
        Assert.Equal(120D, c.Height, 3);
    }

    [Fact]
    public void HtmlFlexColumnWrap_WithAutoHeightRemainsOneNaturalColumn() {
        const string html = """
            <div style="display:flex;flex-direction:column;flex-wrap:wrap;width:100px;row-gap:5px;align-items:flex-start">
              <div id="auto-column-a" style="width:30px;height:20px;background:#ff0000"></div>
              <div id="auto-column-b" style="width:40px;height:30px;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 120D);

        Assert.Equal(0D, FindFlexShape(rendered, "div#auto-column-a").X, 3);
        Assert.Equal(0D, FindFlexShape(rendered, "div#auto-column-a").Y, 3);
        Assert.Equal(0D, FindFlexShape(rendered, "div#auto-column-b").X, 3);
        Assert.Equal(25D, FindFlexShape(rendered, "div#auto-column-b").Y, 3);
    }

    [Fact]
    public void HtmlFlexWrapReverse_ReversesFlexItemCrossStartInsideEachLine() {
        const string html = """
            <div style="display:flex;flex-wrap:wrap-reverse;width:120px;height:80px;align-items:flex-start">
              <div id="row-cross-start" style="width:50px;height:20px;background:#ff0000"></div>
              <div style="width:50px;height:40px;background:#0000ff"></div>
            </div>
            <div style="display:flex;flex-direction:column;flex-wrap:wrap-reverse;width:100px;height:80px;align-items:flex-start">
              <div id="column-cross-start" style="width:30px;height:40px;background:#00ff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 140D);

        Assert.Equal(60D, FindFlexShape(rendered, "div#row-cross-start").Y, 3);
        Assert.Equal(70D, FindFlexShape(rendered, "div#column-cross-start").X, 3);
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
        Assert.Equal(3, rendered.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexValueUnsupported));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
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
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(10, 10));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(40, 10));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"8\" y=\"8\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("<rect x=\"38\" y=\"8\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("FlexPdfMarker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlFlexItems_IncludeAnonymousTextGeneratedContentDisplayContentsAndLinks() {
        const string html = """
            <style>
              #flex::before { content:'Before'; order:-1; width:60px; height:20px; background:#00ff00 }
              #flex::after { content:'After' }
            </style>
            <div id="flex" style="display:flex;width:400px;gap:10px">
              Direct
              <span style="display:contents"><span id="middle" style="width:80px;height:20px;background:#ff0000">Middle</span></span>
            </div>
            <a href="https://example.com/path" style="display:flex">LinkedDirect</a>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 420D);
        IReadOnlyList<HtmlRenderText> texts = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToList();
        HtmlRenderText before = Assert.Single(texts, text => text.Text == "Before");
        HtmlRenderText direct = Assert.Single(texts, text => text.Text == "Direct");
        HtmlRenderText middle = Assert.Single(texts, text => text.Text == "Middle");
        HtmlRenderText after = Assert.Single(texts, text => text.Text == "After");
        HtmlRenderText linked = Assert.Single(texts, text => text.Text == "LinkedDirect");

        Assert.True(before.X < direct.X);
        Assert.True(direct.X < middle.X);
        Assert.True(middle.X < after.X);
        Assert.Equal(60D, FindFlexShape(rendered, "div#flex::before").Width, 3);
        Assert.Equal("https://example.com/path", linked.LinkUri);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexAutoMargins_AbsorbMainAndCrossAxisFreeSpace() {
        HtmlRenderDocument row = RenderFlex("""
            <div style="display:flex;width:300px">
              <div id="row-auto-a" style="margin-left:auto;width:50px;height:20px;background:#ff0000"></div>
              <div id="row-auto-b" style="width:50px;height:20px;background:#0000ff"></div>
            </div>
            """, 320D);
        HtmlRenderDocument rowReverse = RenderFlex("""
            <div style="display:flex;flex-direction:row-reverse;width:300px">
              <div id="row-reverse-auto-a" style="margin-right:auto;width:50px;height:20px;background:#ff0000"></div>
              <div id="row-reverse-auto-b" style="width:50px;height:20px;background:#0000ff"></div>
            </div>
            """, 320D);
        HtmlRenderDocument rowCross = RenderFlex("""
            <div style="display:flex;width:100px;height:100px">
              <div id="row-cross-auto" style="margin-top:auto;width:20px;height:20px;background:#00ff00"></div>
            </div>
            """, 120D);
        HtmlRenderDocument column = RenderFlex("""
            <div style="display:flex;flex-direction:column;width:100px;height:300px">
              <div id="column-auto-a" style="margin-top:auto;height:50px;background:#ff0000"></div>
              <div id="column-auto-b" style="height:50px;background:#0000ff"></div>
            </div>
            """, 120D);
        HtmlRenderDocument columnCross = RenderFlex("""
            <div style="display:flex;flex-direction:column;width:100px;height:40px;align-items:flex-start">
              <div id="column-cross-auto" style="margin-left:auto;width:20px;height:40px;background:#ffff00"></div>
            </div>
            """, 120D);

        Assert.Equal(200D, FindFlexShape(row, "div#row-auto-a").X, 3);
        Assert.Equal(250D, FindFlexShape(row, "div#row-auto-b").X, 3);
        Assert.Equal(50D, FindFlexShape(rowReverse, "div#row-reverse-auto-a").X, 3);
        Assert.Equal(0D, FindFlexShape(rowReverse, "div#row-reverse-auto-b").X, 3);
        Assert.Equal(80D, FindFlexShape(rowCross, "div#row-cross-auto").Y, 3);
        Assert.Equal(200D, FindFlexShape(column, "div#column-auto-a").Y, 3);
        Assert.Equal(250D, FindFlexShape(column, "div#column-auto-b").Y, 3);
        Assert.Equal(80D, FindFlexShape(columnCross, "div#column-cross-auto").X, 3);
    }

    [Fact]
    public void HtmlInlineFlex_ParticipatesAsAnAtomicInlineBox() {
        const string html = """
            <p style="margin:0">Before <a href="https://example.com/inline"><span id="inline" style="display:inline-flex;width:80px;height:20px;gap:10px">
              <span id="inline-a" style="width:20px;height:20px;background:#ff0000"></span>
              <span id="inline-b" style="width:20px;height:20px;background:#0000ff"></span>
            </span></a> After</p>
            """;

        HtmlRenderDocument rendered = RenderFlex(html, 240D);
        HtmlRenderShape a = FindFlexShape(rendered, "span#inline-a");
        HtmlRenderShape b = FindFlexShape(rendered, "span#inline-b");
        HtmlRenderText before = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Before", StringComparison.Ordinal));
        HtmlRenderText after = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("After", StringComparison.Ordinal));

        Assert.True(a.X > before.X);
        Assert.Equal(a.X + 30D, b.X, 3);
        Assert.True(after.X > b.X);
        Assert.Equal(a.Y, b.Y, 3);
        Assert.Contains(rendered.Pages[0].Visuals, visual => visual.Source == "span#inline" && visual.LinkUri == "https://example.com/inline");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending);
    }

    [Fact]
    public void HtmlFlexColumn_Reflows_Percentage_Children_When_Main_Size_Becomes_Definite() {
        HtmlRenderDocument rendered = RenderFlex("""
            <div style="display:flex;flex-direction:column;width:100px;height:40px;align-items:flex-start">
              <div id="item" style="flex:none;width:100px;background:#eeeeee">
                <div id="percent-child" style="height:50%;background:#2563eb">Marker</div>
              </div>
            </div>
            """, 120D);

        HtmlRenderShape item = FindFlexShape(rendered, "div#item");
        HtmlRenderShape child = FindFlexShape(rendered, "div#percent-child");

        Assert.Equal(item.Height * 0.5D, child.Height, 3);
    }

    private static HtmlRenderDocument RenderFlex(string html, double viewportWidth) =>
        HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = viewportWidth,
            Margins = HtmlRenderMargins.All(0D)
        });

    private static HtmlRenderShape FindFlexShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);
}
