using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRelativePosition_PreservesNormalFlowAndUsesLeadingInsets() {
        const string baselineHtml = "<div id='moved' style='height:24px;margin:0;background:#ff0000'>Moved</div>"
            + "<div id='next' style='height:24px;margin:0'>Next</div>";
        const string positionedHtml = "<style>#moved{position:relative;left:10%;right:80px;top:6px;bottom:20px}</style>"
            + "<div id='moved' style='height:24px;margin:0;background:#ff0000'>Moved</div>"
            + "<div id='next' style='height:24px;margin:0'>Next</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument baseline = HtmlRenderEngine.Render(baselineHtml, options);
        HtmlRenderDocument positioned = HtmlRenderEngine.Render(positionedHtml, options);
        HtmlRenderText baselineMoved = FindText(baseline, "Moved");
        HtmlRenderText positionedMoved = FindText(positioned, "Moved");
        HtmlRenderText baselineNext = FindText(baseline, "Next");
        HtmlRenderText positionedNext = FindText(positioned, "Next");

        Assert.Equal(baselineMoved.X + 20D, positionedMoved.X, 3);
        Assert.Equal(baselineMoved.Y + 6D, positionedMoved.Y, 3);
        Assert.Equal(baselineNext.X, positionedNext.X, 3);
        Assert.Equal(baselineNext.Y, positionedNext.Y, 3);
        Assert.Equal(baseline.Pages[0].Height, positioned.Pages[0].Height, 3);
        Assert.DoesNotContain(positioned.Diagnostics.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.PositionInsetUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlRelativePosition_AccumulatesNestedInlineOffsetsWithoutMovingFollowingText() {
        const string baselineHtml = "<p style='margin:0'><span>Outer<span>Inner</span></span><span>Tail</span></p>";
        const string positionedHtml = "<p style='margin:0'><span style='position:relative;left:7px;top:3px'>Outer"
            + "<span style='position:relative;left:5px;top:2px'>Inner</span></span><span>Tail</span></p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument baseline = HtmlRenderEngine.Render(baselineHtml, options);
        HtmlRenderDocument positioned = HtmlRenderEngine.Render(positionedHtml, options);

        AssertPaintOffset(baseline, positioned, "Outer", 7D, 3D);
        AssertPaintOffset(baseline, positioned, "Inner", 12D, 5D);
        AssertPaintOffset(baseline, positioned, "Tail", 0D, 0D);
        Assert.Equal(baseline.Pages[0].Height, positioned.Pages[0].Height, 3);
    }

    [Fact]
    public void HtmlRelativePosition_PaginationUsesNormalFlowCoordinates() {
        string children = string.Concat(Enumerable.Range(1, 6)
            .Select(index => "<div style='height:30px;margin:0'>Marker" + index + "</div>"));
        string baselineHtml = "<section>" + children + "</section>";
        string positionedHtml = "<section style='position:relative;top:40px'>" + children + "</section>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument baseline = HtmlRenderEngine.Render(baselineHtml, options);
        HtmlRenderDocument positioned = HtmlRenderEngine.Render(positionedHtml, options);

        Assert.Equal(baseline.Pages.Count, positioned.Pages.Count);
        for (int index = 1; index <= 6; index++) {
            string marker = "Marker" + index;
            (int BaselinePage, HtmlRenderText BaselineText) = FindTextWithPage(baseline, marker);
            (int PositionedPage, HtmlRenderText PositionedText) = FindTextWithPage(positioned, marker);
            Assert.Equal(BaselinePage, PositionedPage);
            Assert.Equal(BaselineText.X, PositionedText.X, 3);
            Assert.Equal(BaselineText.Y + 40D, PositionedText.Y, 3);
        }
    }

    [Fact]
    public void HtmlRelativePosition_MovesRepeatedTableGroupsWithoutChangingFragments() {
        string rows = string.Concat(Enumerable.Range(0, 18)
            .Select(index => "<tr><td>Row" + index.ToString("D2") + "</td></tr>"));
        string table = "<table><thead><tr><th>HeaderMarker</th></tr></thead>"
            + "<tfoot><tr><td>FooterMarker</td></tr></tfoot><tbody>" + rows + "</tbody></table>";
        string positionedTable = table.Replace("<table>", "<table style='position:relative;left:6px;top:8px'>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(3D, 2D),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(16D)
        };

        HtmlRenderDocument baseline = HtmlRenderEngine.Render(table, options);
        HtmlRenderDocument positioned = HtmlRenderEngine.Render(positionedTable, options);

        Assert.True(baseline.Pages.Count >= 3);
        Assert.Equal(baseline.Pages.Count, positioned.Pages.Count);
        for (int index = 0; index < baseline.Pages.Count; index++) {
            HtmlRenderPage baselinePage = baseline.Pages[index];
            HtmlRenderPage positionedPage = positioned.Pages[index];
            HtmlRenderText baselineHeader = Assert.Single(baselinePage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "HeaderMarker");
            HtmlRenderText positionedHeader = Assert.Single(positionedPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "HeaderMarker");
            HtmlRenderText baselineFooter = Assert.Single(baselinePage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "FooterMarker");
            HtmlRenderText positionedFooter = Assert.Single(positionedPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "FooterMarker");
            Assert.Equal(baselineHeader.X + 6D, positionedHeader.X, 3);
            Assert.Equal(baselineHeader.Y + 8D, positionedHeader.Y, 3);
            Assert.Equal(baselineFooter.X + 6D, positionedFooter.X, 3);
            Assert.Equal(baselineFooter.Y + 8D, positionedFooter.Y, 3);
        }

        Assert.DoesNotContain(positioned.Diagnostics.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed
            || diagnostic.Code == HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed);
    }

    [Fact]
    public void HtmlRelativePosition_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div id='paint' style='position:relative;left:10px;top:10px;width:20px;height:20px;margin:0;background:#ff0000'></div>"
            + "<p style='margin:0'>PositionPdfMarker</p>";
        var options = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 100D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = html.ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);

        Assert.Equal(OfficeColor.White, raster.GetPixel(5, 5));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 15));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"10\" y=\"10\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        string searchablePdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("PositionPdfMarker", searchablePdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlAbsolutePosition_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div style='position:relative;width:100px;height:50px;margin:0'>"
            + "<div id='absolute-paint' style='position:absolute;left:10px;top:10px;width:80px;height:40px;margin:0;background:#ff0000'>AbsolutePdfMarker</div></div>";
        var options = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 100D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = html.ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 50D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);

        Assert.Equal(OfficeColor.White, raster.GetPixel(5, 5));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 15));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"10\" y=\"10\" width=\"80\" height=\"40\"", svg, StringComparison.Ordinal);
        string searchablePdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("AbsolutePdfMarker", searchablePdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlPositioning_DiagnosesUnsupportedModesInsetsAndStacking() {
        const string html = "<div style='position:absolute'>Absolute</div>"
            + "<div style='position:fixed'>Fixed</div>"
            + "<div style='position:sticky'>Sticky</div>"
            + "<div style='position:relative;left:calc(5px + 2px);top:10%;z-index:3'>Relative</div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
        Assert.Equal(2, rendered.Diagnostics.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionInsetUnsupported));
        Assert.Single(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStickyStatic);
        Assert.Single(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionZIndexPending);
        Assert.All(
            new[] {
                HtmlRenderDiagnosticCodes.PositionInsetUnsupported,
                HtmlRenderDiagnosticCodes.PositioningModeUnsupported,
                HtmlRenderDiagnosticCodes.PositionStickyStatic,
                HtmlRenderDiagnosticCodes.PositionZIndexPending
            },
            code => Assert.True(HtmlDiagnosticCatalog.TryGet(code, out _), code));
    }

    [Fact]
    public void HtmlAbsolutePosition_UsesNearestPositionedAncestorAndDoesNotConsumeFlowSpace() {
        const string html = "<section id='host' style='position:relative;width:200px;height:100px;margin:0;background:#eeeeee'>"
            + "<div><div id='absolute' style='position:absolute;left:10%;top:10px;width:20px;height:20px;margin:0;background:#ff0000'>Absolute</div></div>"
            + "<div id='flow' style='width:30px;height:20px;margin:0;background:#0000ff'>Flow</div></section>"
            + "<div id='after' style='width:30px;height:10px;margin:0;background:#00ff00'>After</div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape absolute = FindPositionedShape(rendered, "div#absolute");
        HtmlRenderShape flow = FindPositionedShape(rendered, "div#flow");
        HtmlRenderShape after = FindPositionedShape(rendered, "div#after");
        Assert.Equal(20D, absolute.X, 3);
        Assert.Equal(10D, absolute.Y, 3);
        Assert.Equal(20D, absolute.Width, 3);
        Assert.Equal(20D, absolute.Height, 3);
        Assert.InRange(flow.Y, 0D, 0.02D);
        Assert.Equal(100D, after.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlAbsolutePosition_StretchesBetweenOpposingInsets() {
        const string html = "<div style='position:relative;width:200px;height:100px;margin:0'>"
            + "<div id='stretched' style='position:absolute;left:10px;right:20px;top:10px;bottom:20px;margin:0;background:#ff0000'></div></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 220D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape stretched = FindPositionedShape(rendered, "div#stretched");
        Assert.Equal(10D, stretched.X, 3);
        Assert.Equal(10D, stretched.Y, 3);
        Assert.Equal(170D, stretched.Width, 3);
        Assert.Equal(70D, stretched.Height, 3);
    }

    [Fact]
    public void HtmlAbsolutePosition_ComposesWithFlexAndGridContainers() {
        const string html = "<div style='display:flex;width:200px;height:50px;margin:0'>"
            + "<div id='flex-flow' style='width:20px;height:20px;background:#0000ff'></div>"
            + "<div id='flex-absolute' style='position:absolute;left:100px;top:10px;width:20px;height:20px;background:#ff0000'></div></div>"
            + "<div style='display:grid;grid-template-columns:1fr;width:200px;height:50px;margin:0'>"
            + "<div id='grid-flow' style='height:20px;background:#00ff00'></div>"
            + "<div id='grid-absolute' style='position:absolute;left:100px;top:10px;width:20px;height:20px;background:#ffff00'></div></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 220D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape flexAbsolute = FindPositionedShape(rendered, "div#flex-absolute");
        HtmlRenderShape gridAbsolute = FindPositionedShape(rendered, "div#grid-absolute");
        Assert.Equal(100D, flexAbsolute.X, 3);
        Assert.Equal(10D, flexAbsolute.Y, 3);
        Assert.Equal(100D, gridAbsolute.X, 3);
        Assert.Equal(60D, gridAbsolute.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending
            || diagnostic.Code == HtmlRenderDiagnosticCodes.GridLayoutPending
            || diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlFixedPosition_RepeatsAtViewportCoordinatesOnEveryPage() {
        const string html = "<div id='fixed' style='position:fixed;left:5px;top:5px;width:20px;height:20px;margin:0;background:#ff0000'>Fixed</div>"
            + "<div style='height:60px;margin:0'>One</div><div style='height:60px;margin:0'>Two</div>"
            + "<div style='height:60px;margin:0'>Three</div><div style='height:60px;margin:0'>Four</div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);

        Assert.True(rendered.Pages.Count >= 4);
        foreach (HtmlRenderPage page in rendered.Pages) {
            HtmlRenderShape fixedShape = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#fixed" && shape.Shape.FillColor.HasValue);
            Assert.Equal(5D, fixedShape.X, 3);
            Assert.Equal(5D, fixedShape.Y, 3);
        }
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlStickyPosition_UsesStableDocumentSnapshotAndRemainsInFlow() {
        const string html = "<div id='sticky' style='position:sticky;top:0;width:20px;height:20px;margin:0;background:#ff0000'>Sticky</div>"
            + "<div id='after-sticky' style='width:20px;height:20px;margin:0;background:#0000ff'>After</div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Equal(0D, FindPositionedShape(rendered, "div#sticky").Y, 3);
        Assert.Equal(20D, FindPositionedShape(rendered, "div#after-sticky").Y, 3);
        Assert.Single(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStickyStatic);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlAbsolutePosition_DiagnosesInlineContainingBlockFallbackWithoutConsumingTextFlow() {
        const string baselineHtml = "<p style='margin:0'><span>Before</span><span>After</span></p>";
        const string positionedHtml = "<p style='margin:0'><span style='position:relative'>Before"
            + "<span style='position:absolute;left:5px;top:5px'>Overlay</span></span><span>After</span></p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 160D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument baseline = HtmlRenderEngine.Render(baselineHtml, options);
        HtmlRenderDocument positioned = HtmlRenderEngine.Render(positionedHtml, options);

        HtmlRenderText baselineAfter = FindText(baseline, "After");
        HtmlRenderText positionedAfter = FindText(positioned, "After");
        HtmlRenderText overlay = FindText(positioned, "Overlay");
        Assert.Equal(baselineAfter.X, positionedAfter.X, 3);
        Assert.Equal(baselineAfter.Y, positionedAfter.Y, 3);
        Assert.Equal(5D, overlay.X, 3);
        Assert.Equal(5D, overlay.Y, 3);
        Assert.Single(positioned.Diagnostics.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("containing-block=span", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRelativePosition_ResolvesTrailingInsetsAndExplicitVerticalPercentages() {
        const string baselineHtml = "<div style='height:100px;margin:0'><span>Marker</span></div>";
        const string positionedHtml = "<div style='height:100px;margin:0'><span style='position:relative;right:5px;bottom:10%'>Marker</span></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument baseline = HtmlRenderEngine.Render(baselineHtml, options);
        HtmlRenderDocument positioned = HtmlRenderEngine.Render(positionedHtml, options);

        AssertPaintOffset(baseline, positioned, "Marker", -5D, -10D);
        Assert.DoesNotContain(positioned.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionInsetUnsupported);
    }

    private static void AssertPaintOffset(HtmlRenderDocument baseline, HtmlRenderDocument positioned, string marker, double offsetX, double offsetY) {
        HtmlRenderText baselineText = FindText(baseline, marker);
        HtmlRenderText positionedText = FindText(positioned, marker);
        Assert.Equal(baselineText.X + offsetX, positionedText.X, 3);
        Assert.Equal(baselineText.Y + offsetY, positionedText.Y, 3);
    }

    private static HtmlRenderText FindText(HtmlRenderDocument document, string marker) =>
        document.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Single(text => text.Text == marker);

    private static HtmlRenderShape FindPositionedShape(HtmlRenderDocument document, string source) =>
        Assert.Single(document.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);

    private static (int PageNumber, HtmlRenderText Text) FindTextWithPage(HtmlRenderDocument document, string marker) {
        foreach (HtmlRenderPage page in document.Pages) {
            HtmlRenderText? text = page.Visuals.OfType<HtmlRenderText>().SingleOrDefault(item => item.Text == marker);
            if (text != null) return (page.PageNumber, text);
        }

        throw new Xunit.Sdk.XunitException("Rendered text marker was not found: " + marker);
    }
}
