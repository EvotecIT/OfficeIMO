using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlGrid_ResolvesFixedFractionAndImplicitTracks() {
        const string html = """
            <div style="display:grid;width:300px;grid-template-columns:100px 1fr 2fr;grid-auto-rows:40px;gap:5px 10px">
              <div id="grid-a" style="background:#ff0000"></div>
              <div id="grid-b" style="background:#0000ff"></div>
              <div id="grid-c" style="background:#00ff00"></div>
              <div id="grid-d" style="background:#ffff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 320D);
        HtmlRenderShape a = FindGridShape(rendered, "div#grid-a");
        HtmlRenderShape b = FindGridShape(rendered, "div#grid-b");
        HtmlRenderShape c = FindGridShape(rendered, "div#grid-c");
        HtmlRenderShape d = FindGridShape(rendered, "div#grid-d");

        Assert.Equal(0D, a.X, 3);
        Assert.Equal(100D, a.Width, 3);
        Assert.Equal(110D, b.X, 3);
        Assert.Equal(60D, b.Width, 3);
        Assert.Equal(180D, c.X, 3);
        Assert.Equal(120D, c.Width, 3);
        Assert.Equal(0D, d.X, 3);
        Assert.Equal(45D, d.Y, 3);
        Assert.All(new[] { a, b, c, d }, shape => Assert.Equal(40D, shape.Height, 3));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridLayoutPending);
    }

    [Fact]
    public void HtmlGrid_HonorsNumericPlacementAndSpans() {
        const string html = """
            <div style="display:grid;width:300px;grid-template-columns:repeat(3,1fr);grid-auto-rows:auto;gap:5px 10px">
              <div id="placed-a" style="grid-column:2 / span 2;grid-row:1;height:30px;background:#ff0000"></div>
              <div id="placed-b" style="grid-area:2 / 1 / 3 / 3;height:40px;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 320D);
        HtmlRenderShape a = FindGridShape(rendered, "div#placed-a");
        HtmlRenderShape b = FindGridShape(rendered, "div#placed-b");

        Assert.Equal(310D / 3D, a.X, 3);
        Assert.Equal(590D / 3D, a.Width, 3);
        Assert.Equal(0D, a.Y, 3);
        Assert.Equal(0D, b.X, 3);
        Assert.Equal(590D / 3D, b.Width, 3);
        Assert.Equal(35D, b.Y, 3);
    }

    [Fact]
    public void HtmlGrid_ResolvesRepeatMinmaxAndDenseAutoPlacement() {
        const string html = """
            <div style="display:grid;width:200px;grid-template-columns:minmax(50px,1fr) 1fr;grid-auto-rows:20px">
              <div id="minmax-equal-a" style="background:#ff0000"></div><div id="minmax-equal-b" style="background:#0000ff"></div>
            </div>
            <div style="display:grid;width:200px;grid-template-columns:minmax(150px,1fr) 1fr;grid-auto-rows:20px">
              <div id="minmax-frozen-a" style="background:#00ff00"></div><div id="minmax-frozen-b" style="background:#ffff00"></div>
            </div>
            <div style="display:grid;width:300px;grid-template-columns:repeat(3,1fr);grid-auto-flow:row dense;grid-auto-rows:20px">
              <div id="dense-a" style="grid-column:span 2;background:#ff0000"></div>
              <div id="dense-b" style="grid-column:span 2;background:#0000ff"></div>
              <div id="dense-c" style="background:#00ff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 320D);

        Assert.Equal(100D, FindGridShape(rendered, "div#minmax-equal-a").Width, 3);
        Assert.Equal(100D, FindGridShape(rendered, "div#minmax-equal-b").Width, 3);
        Assert.Equal(150D, FindGridShape(rendered, "div#minmax-frozen-a").Width, 3);
        Assert.Equal(50D, FindGridShape(rendered, "div#minmax-frozen-b").Width, 3);
        HtmlRenderShape denseA = FindGridShape(rendered, "div#dense-a");
        HtmlRenderShape denseB = FindGridShape(rendered, "div#dense-b");
        HtmlRenderShape denseC = FindGridShape(rendered, "div#dense-c");
        Assert.Equal(denseA.Y, denseC.Y, 3);
        Assert.True(denseB.Y > denseA.Y);
        Assert.Equal(200D, denseC.X, 3);
    }

    [Fact]
    public void HtmlGridItems_IncludeAnonymousGeneratedAndDisplayContentsContent() {
        const string html = """
            <style>
              #grid-items::before { content:'Before'; background:#00ff00 }
              #grid-items::after { content:'After' }
            </style>
            <div id="grid-items" style="display:grid;width:240px;grid-template-columns:repeat(2,1fr);grid-auto-rows:30px;gap:5px">
              Direct
              <span style="display:contents"><span id="grid-middle" style="background:#ff0000">Middle</span></span>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 260D);
        IReadOnlyList<HtmlRenderText> texts = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToList();

        Assert.Single(texts, text => text.Text == "Before");
        Assert.Single(texts, text => text.Text == "Direct");
        Assert.Single(texts, text => text.Text == "Middle");
        Assert.Single(texts, text => text.Text == "After");
        Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#grid-items::before" && shape.Shape.FillColor.HasValue);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridLayoutPending);
    }

    [Fact]
    public void HtmlGrid_MapsNamedTemplateAreasToRectangularCells() {
        const string html = """
            <div style='display:grid;width:200px;grid-template-areas:"header header" "side main";grid-template-columns:80px 1fr;grid-auto-rows:30px'>
              <div id="area-header" style="grid-area:header;background:#ff0000"></div>
              <div id="area-side" style="grid-area:side;background:#0000ff"></div>
              <div id="area-main" style="grid-area:main;background:#00ff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 220D);
        HtmlRenderShape header = FindGridShape(rendered, "div#area-header");
        HtmlRenderShape side = FindGridShape(rendered, "div#area-side");
        HtmlRenderShape main = FindGridShape(rendered, "div#area-main");

        Assert.Equal(0D, header.X, 3);
        Assert.Equal(200D, header.Width, 3);
        Assert.Equal(0D, header.Y, 3);
        Assert.Equal(0D, side.X, 3);
        Assert.Equal(80D, side.Width, 3);
        Assert.Equal(30D, side.Y, 3);
        Assert.Equal(80D, main.X, 3);
        Assert.Equal(120D, main.Width, 3);
        Assert.Equal(30D, main.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Fact]
    public void HtmlGrid_ColumnAutoFlowFillsRowsBeforeCreatingImplicitColumns() {
        const string html = """
            <div style="display:grid;width:200px;grid-template-rows:repeat(2,30px);grid-auto-columns:50px;grid-auto-flow:column;gap:5px 10px;justify-content:start">
              <div id="flow-column-a" style="background:#ff0000"></div>
              <div id="flow-column-b" style="background:#0000ff"></div>
              <div id="flow-column-c" style="background:#00ff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 220D);

        Assert.Equal(0D, FindGridShape(rendered, "div#flow-column-a").X, 3);
        Assert.Equal(0D, FindGridShape(rendered, "div#flow-column-a").Y, 3);
        Assert.Equal(0D, FindGridShape(rendered, "div#flow-column-b").X, 3);
        Assert.Equal(35D, FindGridShape(rendered, "div#flow-column-b").Y, 3);
        Assert.Equal(60D, FindGridShape(rendered, "div#flow-column-c").X, 3);
        Assert.Equal(0D, FindGridShape(rendered, "div#flow-column-c").Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Fact]
    public void HtmlGrid_ResolvesNamedTemplateLines() {
        const string html = """
            <div style="display:grid;width:200px;grid-template-columns:[side-start] 80px [side-end main-start] 1fr [main-end];grid-template-rows:[top] 30px [bottom]">
              <div id="named-side" style="grid-column:side-start / side-end;grid-row:top / bottom;background:#ff0000"></div>
              <div id="named-main" style="grid-column:main-start / main-end;grid-row:top / bottom;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 220D);

        Assert.Equal(0D, FindGridShape(rendered, "div#named-side").X, 3);
        Assert.Equal(80D, FindGridShape(rendered, "div#named-side").Width, 3);
        Assert.Equal(80D, FindGridShape(rendered, "div#named-main").X, 3);
        Assert.Equal(120D, FindGridShape(rendered, "div#named-main").Width, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Fact]
    public void HtmlGrid_DistinguishesResponsiveAutoFitAndAutoFillTracks() {
        const string html = """
            <div style="display:grid;width:450px;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));column-gap:10px;grid-auto-rows:20px">
              <div id="auto-fit-a" style="background:#ff0000"></div><div id="auto-fit-b" style="background:#0000ff"></div>
            </div>
            <div style="display:grid;width:450px;grid-template-columns:repeat(auto-fill,minmax(100px,1fr));column-gap:10px;grid-auto-rows:20px">
              <div id="auto-fill-a" style="background:#00ff00"></div><div id="auto-fill-b" style="background:#ffff00"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 470D);

        Assert.Equal(220D, FindGridShape(rendered, "div#auto-fit-a").Width, 3);
        Assert.Equal(230D, FindGridShape(rendered, "div#auto-fit-b").X, 3);
        Assert.Equal(105D, FindGridShape(rendered, "div#auto-fill-a").Width, 3);
        Assert.Equal(115D, FindGridShape(rendered, "div#auto-fill-b").X, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Fact]
    public void HtmlGrid_AppliesContainerAndItemAlignment() {
        const string html = """
            <div style="display:grid;width:200px;height:100px;grid-template-columns:repeat(2,1fr);grid-template-rows:100px;justify-items:center;align-items:end">
              <div id="aligned-a" style="width:20px;height:30px;background:#ff0000"></div>
              <div id="aligned-b" style="width:20px;height:30px;justify-self:end;align-self:center;background:#0000ff"></div>
            </div>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 220D);
        HtmlRenderShape a = FindGridShape(rendered, "div#aligned-a");
        HtmlRenderShape b = FindGridShape(rendered, "div#aligned-b");

        Assert.Equal(40D, a.X, 3);
        Assert.Equal(70D, a.Y, 3);
        Assert.Equal(180D, b.X, 3);
        Assert.Equal(35D, b.Y, 3);
    }

    [Fact]
    public void HtmlGrid_PaginatesOnlyAtUnspannedRowBoundaries() {
        const string html = """
            <div style="height:20px;margin:0">Before</div>
            <div style="display:grid;width:100px;grid-template-columns:1fr;grid-auto-rows:40px">
              <div id="grid-page-one" style="background:#ff0000">One</div>
              <div id="grid-page-two" style="background:#0000ff">Two</div>
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
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#grid-page-one");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#grid-page-two");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#grid-page-two");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlGrid_PaginatesInsideOneOversizedItemAtNestedBlockBoundaries() {
        const string html = """
            <div id="grid" style="display:grid;grid-template-columns:100px">
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
    public void HtmlGrid_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = """
            <div style="display:grid;width:50px;grid-template-columns:20px 20px;grid-template-rows:20px;column-gap:10px">
              <div style="background:#ff0000"></div>
              <div style="background:#0000ff"></div>
            </div>
            <p style="margin:0">GridPdfMarker</p>
            """;
        var options = new HtmlRenderOptions { ViewportWidth = 80D, Margins = HtmlRenderMargins.All(8D) };

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
        Assert.Contains("GridPdfMarker", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlInlineGrid_ParticipatesAsAnAtomicLinkedInlineBox() {
        const string html = """
            <p style="margin:0">Before <a href="https://example.com/grid"><span id="inline-grid" style="display:inline-grid;width:80px;height:20px;grid-template-columns:20px 20px;column-gap:10px">
              <span id="inline-grid-a" style="background:#ff0000"></span>
              <span id="inline-grid-b" style="background:#0000ff"></span>
            </span></a> After</p>
            """;

        HtmlRenderDocument rendered = RenderGrid(html, 240D);
        HtmlRenderShape a = FindGridShape(rendered, "span#inline-grid-a");
        HtmlRenderShape b = FindGridShape(rendered, "span#inline-grid-b");
        HtmlRenderText before = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Before", StringComparison.Ordinal));
        HtmlRenderText after = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("After", StringComparison.Ordinal));

        Assert.True(a.X > before.X);
        Assert.Equal(a.X + 30D, b.X, 3);
        Assert.True(after.X > b.X);
        Assert.Contains(rendered.Pages[0].Visuals, visual => visual.Source == "span#inline-grid" && visual.LinkUri == "https://example.com/grid");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridLayoutPending);
    }

    [Fact]
    public void HtmlGrid_DiagnosesUnsupportedValuesAndBoundsTrackExpansion() {
        const string html = """
            <div style="display:grid;width:200px;grid-template-columns:subgrid 1fr;grid-auto-flow:sideways">
              <div style="grid-column-start:named">One</div><div>Two</div>
            </div>
            """;
        HtmlRenderDocument rendered = RenderGrid(html, 220D);

        Assert.Equal(3, rendered.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridLayoutPending);

        var options = new HtmlRenderOptions { ViewportWidth = 220D, Margins = HtmlRenderMargins.All(0D), MaxGridTracks = 2 };
        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render("<div style='display:grid;grid-template-columns:repeat(3,1fr)'><span>A</span></div>", options));
        Assert.Equal(HtmlRenderDiagnosticCodes.GridTrackLimitExceeded, exception.Code);
    }

    [Fact]
    public void HtmlGrid_BoundsAuthoredLineNumbersBeforeIntegerArithmetic() {
        var options = new HtmlRenderOptions { MaxGridTracks = 16 };

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render("<div style='display:grid'><span style='grid-column:2147483647 / span 2'>A</span></div>", options));

        Assert.Equal(HtmlRenderDiagnosticCodes.GridTrackLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxGridTracks), exception.LimitSource);
    }

    [Fact]
    public void HtmlGrid_ColumnFlowAcceptsItemSpansBeyondExplicitRows() {
        const string html = "<div style='display:grid;grid-auto-flow:column;grid-template-rows:20px;grid-auto-rows:20px'>"
            + "<span id='tall' style='grid-row:span 2;background:red'>A</span></div>";

        HtmlRenderDocument rendered = RenderGrid(html, 100D);

        Assert.Equal(40D, FindGridShape(rendered, "span#tall").Height, 3);
    }

    [Fact]
    public void HtmlGrid_BoundsNestedRepeatFunctionDepth() {
        string tracks = "1px";
        for (int index = 0; index < 8; index++) tracks = "repeat(auto-fit," + tracks + ")";

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render("<div style='display:grid;grid-template-columns:" + tracks + "'><span>A</span></div>",
                new HtmlRenderOptions { MaxLayoutDepth = 4 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.DepthLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutDepth), exception.LimitSource);
    }

    [Fact]
    public void HtmlGrid_Charges_Dense_Auto_Placement_Probes_To_The_Layout_Budget() {
        var html = new StringBuilder("<div style='display:grid;grid-template-columns:repeat(32,1fr);grid-auto-flow:row dense'>");
        for (int index = 0; index < 256; index++) html.Append("<span>").Append(index).Append("</span>");
        html.Append("</div>");

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(html.ToString(), new HtmlRenderOptions { MaxLayoutOperations = 4_000 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
    }

    [Fact]
    public void HtmlGrid_HandlesManyItemsSharingOneRowWithoutQuadraticOccupancyScans() {
        const int itemCount = 2048;
        var html = new StringBuilder("<div style='display:grid;grid-template-columns:1fr;grid-template-rows:20px'>");
        for (int index = 0; index < itemCount; index++) {
            html.Append("<span style='grid-row:1;grid-column:1'>x</span>");
        }
        html.Append("</div>");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            html.ToString(),
            new HtmlRenderOptions {
                ViewportWidth = 100D,
                Margins = HtmlRenderMargins.All(0D),
                MaxLayoutOperations = 100_000
            });

        Assert.Single(rendered.Pages);
        Assert.DoesNotContain(
            rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded);
    }

    private static HtmlRenderDocument RenderGrid(string html, double viewportWidth) =>
        HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions { ViewportWidth = viewportWidth, Margins = HtmlRenderMargins.All(0D) });

    private static HtmlRenderShape FindGridShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);
}
