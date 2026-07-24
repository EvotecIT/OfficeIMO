using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlColumns_BalancesFixedBlocksAcrossRequestedColumns() {
        const string html = "<div style='width:100px;margin:0;column-count:2;column-gap:20px'>"
            + "<div id='column-a' style='height:20px;background:#ff0000'></div>"
            + "<div id='column-b' style='height:20px;background:#00ff00'></div>"
            + "<div id='column-c' style='height:20px;background:#0000ff'></div>"
            + "<div id='column-d' style='height:20px;background:#ffff00'></div></div>";

        HtmlRenderDocument rendered = RenderColumns(html, 120D);
        HtmlRenderShape a = FindColumnShape(rendered, "div#column-a");
        HtmlRenderShape b = FindColumnShape(rendered, "div#column-b");
        HtmlRenderShape c = FindColumnShape(rendered, "div#column-c");
        HtmlRenderShape d = FindColumnShape(rendered, "div#column-d");

        Assert.Equal(0D, a.X, 3);
        Assert.Equal(0D, b.X, 3);
        Assert.Equal(60D, c.X, 3);
        Assert.Equal(60D, d.X, 3);
        Assert.Equal(0D, a.Y, 3);
        Assert.Equal(20D, b.Y, 3);
        Assert.Equal(0D, c.Y, 3);
        Assert.Equal(20D, d.Y, 3);
        Assert.All(new[] { a, b, c, d }, shape => Assert.Equal(40D, shape.Width, 3));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported);
    }

    [Fact]
    public void HtmlColumns_ResolvesColumnWidthAndNormalGap() {
        const string html = "<div style='width:240px;margin:0;column-width:60px'>"
            + "<div id='width-a' style='height:20px;background:#ff0000'></div>"
            + "<div id='width-b' style='height:20px;background:#00ff00'></div>"
            + "<div id='width-c' style='height:20px;background:#0000ff'></div></div>";

        HtmlRenderDocument rendered = RenderColumns(html, 260D);
        HtmlRenderShape a = FindColumnShape(rendered, "div#width-a");
        HtmlRenderShape b = FindColumnShape(rendered, "div#width-b");
        HtmlRenderShape c = FindColumnShape(rendered, "div#width-c");

        Assert.Equal(0D, a.X, 3);
        Assert.Equal(256D / 3D, b.X, 3);
        Assert.Equal(512D / 3D, c.X, 3);
        Assert.All(new[] { a, b, c }, shape => Assert.Equal(208D / 3D, shape.Width, 3));
    }

    [Fact]
    public void HtmlColumns_ColumnFillAutoUsesDeclaredHeightBeforeAdvancing() {
        const string html = "<div style='width:100px;height:40px;margin:0;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<div id='auto-a' style='height:20px;background:#ff0000'></div>"
            + "<div id='auto-b' style='height:20px;background:#00ff00'></div>"
            + "<div id='auto-c' style='height:20px;background:#0000ff'></div></div>";

        HtmlRenderDocument rendered = RenderColumns(html, 120D);
        HtmlRenderShape a = FindColumnShape(rendered, "div#auto-a");
        HtmlRenderShape b = FindColumnShape(rendered, "div#auto-b");
        HtmlRenderShape c = FindColumnShape(rendered, "div#auto-c");

        Assert.Equal(0D, a.X, 3);
        Assert.Equal(0D, b.X, 3);
        Assert.Equal(60D, c.X, 3);
        Assert.Equal(0D, c.Y, 3);
    }

    [Fact]
    public void HtmlColumns_ColumnSpanAllSeparatesBalancedColumnSets() {
        const string html = "<div id='spanning-columns' style='width:100px;margin:0;column-count:2;column-gap:20px'>"
            + "<div id='before-a' style='height:20px;background:#ff0000'></div>"
            + "<div id='before-b' style='height:20px;background:#00ff00'></div>"
            + "<div id='column-spanner' style='column-span:all;height:10px;background:#000000'></div>"
            + "<div id='after-a' style='height:20px;background:#0000ff'></div>"
            + "<div id='after-b' style='height:20px;background:#ffff00'></div></div>";

        HtmlRenderDocument rendered = RenderColumns(html, 120D);
        HtmlRenderShape beforeA = FindColumnShape(rendered, "div#before-a");
        HtmlRenderShape beforeB = FindColumnShape(rendered, "div#before-b");
        HtmlRenderShape spanner = FindColumnShape(rendered, "div#column-spanner");
        HtmlRenderShape afterA = FindColumnShape(rendered, "div#after-a");
        HtmlRenderShape afterB = FindColumnShape(rendered, "div#after-b");

        Assert.Equal(0D, beforeA.X, 3);
        Assert.Equal(60D, beforeB.X, 3);
        Assert.Equal(0D, spanner.X, 3);
        Assert.Equal(100D, spanner.Width, 3);
        Assert.Equal(20D, spanner.Y, 3);
        Assert.Equal(0D, afterA.X, 3);
        Assert.Equal(60D, afterB.X, 3);
        Assert.Equal(30D, afterA.Y, 3);
        Assert.Equal(30D, afterB.Y, 3);
    }

    [Fact]
    public void HtmlColumns_SpanningFlowPreservesGeneratedContentAndDocumentOrder() {
        const string html = "<style>#generated-columns::before{content:'Before'}#generated-columns::after{content:'After'}</style>"
            + "<div id='generated-columns' style='width:120px;margin:0;column-count:2'>Start"
            + "<div style='column-span:all'>Spanner</div>End</div>";

        HtmlRenderDocument rendered = RenderColumns(html, 140D);
        string text = string.Join("|", rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
            .OrderBy(item => item.Y).ThenBy(item => item.X).Select(item => item.Text));

        Assert.Contains("Before", text, StringComparison.Ordinal);
        Assert.Contains("Start", text, StringComparison.Ordinal);
        Assert.Contains("Spanner", text, StringComparison.Ordinal);
        Assert.Contains("End", text, StringComparison.Ordinal);
        Assert.Contains("After", text, StringComparison.Ordinal);
        Assert.True(text.IndexOf("Before", StringComparison.Ordinal) < text.IndexOf("Spanner", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Spanner", StringComparison.Ordinal) < text.IndexOf("After", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlColumns_ColumnRuleUsesSharedVectorPaint() {
        const string html = "<div id='rule-columns' style='width:100px;margin:0;column-count:2;column-gap:20px;column-rule:4px dashed #ff00ff'>"
            + "<div style='height:20px'></div><div style='height:20px'></div></div>";

        HtmlRenderDocument rendered = RenderColumns(html, 120D);
        HtmlRenderShape rule = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#rule-columns::column-rule");

        Assert.Equal(50D, rule.X, 3);
        Assert.Equal(20D, rule.Height, 3);
        Assert.Equal(4D, rule.Shape.StrokeWidth, 3);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 255), rule.Shape.StrokeColor);
        Assert.Equal(OfficeStrokeDashStyle.Dash, rule.Shape.StrokeDashStyle);
    }

    [Fact]
    public void HtmlColumns_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div style='width:100px;height:20px;margin:0;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<div style='height:20px;background:#ff0000;font-size:10px;line-height:10px'>ColPdf</div>"
            + "<div style='height:20px;background:#0000ff;font-size:10px;line-height:10px'>Second</div></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(120D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, raster.GetPixel(38, 18));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(98, 18));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"0\" y=\"0\" width=\"40\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("<rect x=\"60\" y=\"0\" width=\"40\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("ColPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlColumns_PaginatesFragmentsWithoutDroppingContent() {
        const string html = "<div style='width:100px;margin:0;column-count:2;column-gap:20px;column-rule:2px solid black'>"
            + "<div id='paged-column-a' style='height:20px;background:#ff0000'>One</div>"
            + "<div style='height:20px'></div>"
            + "<div style='height:20px'></div>"
            + "<div style='height:20px'></div>"
            + "<div style='height:20px'></div>"
            + "<div id='paged-column-b' style='height:20px;background:#0000ff'>Six</div></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(120D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text.Contains("One", StringComparison.Ordinal));
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text.Contains("Six", StringComparison.Ordinal));
        Assert.All(rendered.Pages, page => {
            HtmlRenderClipGroup fragment = Assert.Single(page.Visuals.OfType<HtmlRenderClipGroup>(), group => group.Source == "div::column-rule");
            HtmlRenderShape rule = Assert.Single(fragment.Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div::column-rule");
            Assert.Equal(OfficeShapeKind.Line, rule.Shape.Kind);
            Assert.True(rule.Shape.Height > fragment.ClipHeight);
        });
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlColumns_DiagnosesUnsupportedValuesAndBoundsGeneration() {
        const string html = "<div id='invalid-columns' style='columns:wide 2;column-fill:spread;column-span:some;column-rule:wavy'>Text</div>";
        HtmlRenderDocument rendered = RenderColumns(html, 120D);

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported);
        Assert.Equal("div#invalid-columns", diagnostic.Source);
        Assert.Contains("wide", diagnostic.Detail);
        Assert.Contains("column-fill=spread", diagnostic.Detail);
        Assert.Contains("column-span=some", diagnostic.Detail);
        Assert.Contains("wavy", diagnostic.Detail);
        Assert.Contains(HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded, HtmlRenderDiagnosticCodes.All);
        Assert.Contains(HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded, out _));
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(column-count:2)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(column-fill:balance)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(column-span:all)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(column-width:10em)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(columns:10em 2)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(column-rule:2px dashed red)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(column-count:0)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(column-fill:spread)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(column-span:some)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(column-width:wide)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(column-width:50%)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(column-width:10)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(columns:wide 2)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(column-rule:wavy)"));

        var options = new HtmlRenderOptions { ViewportWidth = 120D, Margins = HtmlRenderMargins.All(0D), MaxColumnCount = 2 };
        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render("<div style='column-count:3'>Text</div>", options));
        Assert.Equal(HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxColumnCount), exception.LimitSource);
        Assert.Equal(3L, exception.Actual);
        Assert.Equal(2L, exception.Limit);
    }

    [Fact]
    public void HtmlColumns_BalancingProbesMayTemporarilyExceedTheFinalColumnLimit() {
        const string html = """
            <div style='column-count:2;column-fill:balance'>
              <div style='height:40px;break-inside:avoid'>One</div>
              <div style='height:40px;break-inside:avoid'>Two</div>
              <div style='height:40px;break-inside:avoid'>Three</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 160D,
            Margins = HtmlRenderMargins.All(0D),
            MaxColumnCount = 2
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text.Contains("One", StringComparison.Ordinal));
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text.Contains("Three", StringComparison.Ordinal));
    }

    private static HtmlRenderDocument RenderColumns(string html, double viewportWidth) =>
        HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions { ViewportWidth = viewportWidth, Margins = HtmlRenderMargins.All(0D) });

    private static HtmlRenderShape FindColumnShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);
}
