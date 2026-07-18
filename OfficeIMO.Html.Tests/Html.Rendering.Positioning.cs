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

        HtmlRenderDocument baseline = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(baselineHtml), options);
        HtmlRenderDocument positioned = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(positionedHtml), options);
        HtmlRenderText baselineMoved = FindText(baseline, "Moved");
        HtmlRenderText positionedMoved = FindText(positioned, "Moved");
        HtmlRenderText baselineNext = FindText(baseline, "Next");
        HtmlRenderText positionedNext = FindText(positioned, "Next");

        Assert.Equal(baselineMoved.X + 20D, positionedMoved.X, 3);
        Assert.Equal(baselineMoved.Y + 6D, positionedMoved.Y, 3);
        Assert.Equal(baselineNext.X, positionedNext.X, 3);
        Assert.Equal(baselineNext.Y, positionedNext.Y, 3);
        Assert.Equal(baseline.Pages[0].Height, positioned.Pages[0].Height, 3);
        Assert.DoesNotContain(positioned.Diagnostics, diagnostic =>
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

        HtmlRenderDocument baseline = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(baselineHtml), options);
        HtmlRenderDocument positioned = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(positionedHtml), options);

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

        HtmlRenderDocument baseline = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(baselineHtml), options);
        HtmlRenderDocument positioned = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(positionedHtml), options);

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

        HtmlRenderDocument baseline = HtmlRenderTestDriver.Render(table, options);
        HtmlRenderDocument positioned = HtmlRenderTestDriver.Render(positionedTable, options);

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

        Assert.DoesNotContain(positioned.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed
            || diagnostic.Code == HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed);
    }

    [Fact]
    public void HtmlRelativePosition_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div id='paint' style='position:relative;left:10px;top:10px;width:20px;height:20px;margin:0;background:#ff0000'></div>"
            + "<p style='margin:0'>PositionPdfMarker</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 100D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(OfficeColor.White, raster.GetPixel(5, 5));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 15));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"10\" y=\"10\" width=\"20\" height=\"20\"", svg, StringComparison.Ordinal);
        string searchablePdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("PositionPdfMarker", searchablePdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlAbsolutePosition_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div style='position:relative;width:100px;height:50px;margin:0'>"
            + "<div id='absolute-paint' style='position:absolute;left:10px;top:10px;width:80px;height:40px;margin:0;background:#ff0000'>AbsolutePdfMarker</div></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 100D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 50D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(OfficeColor.White, raster.GetPixel(5, 5));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(85, 45));
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<rect x=\"10\" y=\"10\" width=\"80\" height=\"40\"", svg, StringComparison.Ordinal);
        string searchablePdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("AbsolutePdfMarker", searchablePdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlPositioning_DiagnosesUnsupportedModesInsetsAndStacking() {
        const string html = "<div style='position:absolute'>Absolute</div>"
            + "<div style='position:fixed'>Fixed</div>"
            + "<div style='position:sticky'>Sticky</div>"
            + "<div style='position:relative;left:calc(5px + 2px);top:10%;z-index:3'>Relative</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
        Assert.Equal(2, rendered.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionInsetUnsupported));
        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStickyStatic);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionZIndexPending);
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlAbsolutePosition_StretchesBetweenOpposingInsets() {
        const string html = "<div style='position:relative;width:200px;height:100px;margin:0'>"
            + "<div id='stretched' style='position:absolute;left:10px;right:20px;top:10px;bottom:20px;margin:0;background:#ff0000'></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 220D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape flexAbsolute = FindPositionedShape(rendered, "div#flex-absolute");
        HtmlRenderShape gridAbsolute = FindPositionedShape(rendered, "div#grid-absolute");
        Assert.Equal(100D, flexAbsolute.X, 3);
        Assert.Equal(10D, flexAbsolute.Y, 3);
        Assert.Equal(100D, gridAbsolute.X, 3);
        Assert.Equal(60D, gridAbsolute.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.True(rendered.Pages.Count >= 4);
        foreach (HtmlRenderPage page in rendered.Pages) {
            HtmlRenderShape fixedShape = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#fixed" && shape.Shape.FillColor.HasValue);
            Assert.Equal(5D, fixedShape.X, 3);
            Assert.Equal(5D, fixedShape.Y, 3);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlStickyPosition_UsesStableDocumentSnapshotAndRemainsInFlow() {
        const string html = "<div id='sticky' style='position:sticky;z-index:2;top:0;width:20px;height:20px;margin:0;background:#ff0000'>Sticky</div>"
            + "<div id='after-sticky' style='width:20px;height:20px;margin:0;background:#0000ff'>After</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Equal(0D, FindPositionedShape(rendered, "div#sticky").Y, 3);
        Assert.Equal(20D, FindPositionedShape(rendered, "div#after-sticky").Y, 3);
        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStickyStatic);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionZIndexPending);
    }

    [Fact]
    public void HtmlAbsolutePosition_UsesInlineContainingBlockWithoutConsumingTextFlow() {
        const string baselineHtml = "<p style='margin:0'><span>Before</span><span>After</span></p>";
        const string positionedHtml = "<p style='margin:0'><span style='position:relative'>Before"
            + "<span style='position:absolute;left:5px;top:5px'>Overlay</span></span><span>After</span></p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 160D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument baseline = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(baselineHtml), options);
        HtmlRenderDocument positioned = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(positionedHtml), options);

        HtmlRenderText baselineAfter = FindText(baseline, "After");
        HtmlRenderText positionedAfter = FindText(positioned, "After");
        HtmlRenderText overlay = FindText(positioned, "Overlay");
        Assert.Equal(baselineAfter.X, positionedAfter.X, 3);
        Assert.Equal(baselineAfter.Y, positionedAfter.Y, 3);
        Assert.Equal(5D, overlay.X, 3);
        Assert.Equal(5D, overlay.Y, 3);
        Assert.DoesNotContain(positioned.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlAbsolutePosition_OrdersNegativeFlowAutoAndPositiveStackingBands() {
        const string html = "<div id='stack' style='position:relative;width:40px;height:40px;margin:0;background:#eeeeee'>"
            + "<div id='negative' style='position:absolute;z-index:-2;left:0;top:0;width:40px;height:40px;background:#ff0000'></div>"
            + "<div id='flow-layer' style='width:40px;height:40px;margin:0;background:#00ff00'></div>"
            + "<div id='auto-layer' style='position:absolute;left:0;top:0;width:40px;height:40px;background:#0000ff'></div>"
            + "<div id='positive-two' style='position:absolute;z-index:2;left:0;top:0;width:40px;height:40px;background:#ffff00'></div>"
            + "<div id='positive-one' style='position:absolute;z-index:1;left:0;top:0;width:40px;height:40px;background:#ff00ff'></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] paintSources = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "div#stack"
                || shape.Source == "div#negative"
                || shape.Source == "div#flow-layer"
                || shape.Source == "div#auto-layer"
                || shape.Source == "div#positive-one"
                || shape.Source == "div#positive-two")
            .Select(shape => shape.Source!)
            .ToArray();
        Assert.Equal(new[] { "div#stack", "div#negative", "div#flow-layer", "div#auto-layer", "div#positive-one", "div#positive-two" }, paintSources);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Yellow, raster.GetPixel(20, 20));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionZIndexPending);
    }

    [Fact]
    public void HtmlAbsolutePosition_KeepsDescendantsInsideParentStackingContext() {
        const string html = "<div style='position:relative;width:40px;height:40px;margin:0'>"
            + "<div id='sibling-context' style='position:absolute;z-index:5;left:0;top:0;width:40px;height:40px;background:#0000ff'></div>"
            + "<div id='parent-context' style='position:absolute;z-index:10;left:0;top:0;width:40px;height:40px;background:#ff0000'>"
            + "<div id='nested-negative' style='position:absolute;z-index:-100;left:0;top:0;width:40px;height:40px;background:#00ff00'></div></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] paintSources = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "div#sibling-context" || shape.Source == "div#parent-context" || shape.Source == "div#nested-negative")
            .Select(shape => shape.Source!)
            .ToArray();
        Assert.Equal(new[] { "div#sibling-context", "div#parent-context", "div#nested-negative" }, paintSources);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0x00, 0xFF, 0x00), raster.GetPixel(20, 20));
    }

    [Fact]
    public void HtmlPositioning_OrdersRootAbsoluteAndFixedLayersAroundDocumentFlow() {
        const string html = "<div id='root-positive' style='position:absolute;z-index:2;left:0;top:0;width:40px;height:40px;background:#ffff00'></div>"
            + "<div id='fixed-one' style='position:fixed;z-index:1;left:0;top:0;width:40px;height:40px;background:#0000ff'></div>"
            + "<div id='root-negative' style='position:absolute;z-index:-1;left:0;top:0;width:40px;height:40px;background:#ff0000'></div>"
            + "<div id='document-flow' style='width:40px;height:40px;margin:0;background:#00ff00'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] paintSources = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "div#root-positive"
                || shape.Source == "div#fixed-one"
                || shape.Source == "div#root-negative"
                || shape.Source == "div#document-flow")
            .Select(shape => shape.Source!)
            .ToArray();
        Assert.Equal(new[] { "div#root-negative", "div#document-flow", "div#fixed-one", "div#root-positive" }, paintSources);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Yellow, raster.GetPixel(20, 20));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionZIndexPending);
    }

    [Fact]
    public void HtmlAbsolutePosition_UsesBlockStaticPositionWhenInsetsAreAuto() {
        const string html = "<div id='before-auto' style='height:30px;margin:0;background:#0000ff'></div>"
            + "<div id='auto-positioned' style='position:absolute;width:20px;height:20px;margin:0;background:#ff0000'></div>"
            + "<div id='after-auto' style='height:20px;margin:0;background:#00ff00'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape positioned = FindPositionedShape(rendered, "div#auto-positioned");
        Assert.Equal(0D, positioned.X, 3);
        Assert.Equal(30D, positioned.Y, 3);
        Assert.Equal(30D, FindPositionedShape(rendered, "div#after-auto").Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlAbsolutePosition_AccumulatesNestedBlockStaticPositionToContainingPaddingBox() {
        const string html = "<div style='position:relative;width:100px;height:100px;padding:10px;margin:0'>"
            + "<div style='margin:7px 0 0 5px;border:2px solid #000;padding:3px'>"
            + "<div style='height:20px;margin:0'></div>"
            + "<div id='nested-auto' style='position:absolute;width:10px;height:10px;margin:0;background:#ff0000'></div></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape positioned = FindPositionedShape(rendered, "div#nested-auto");
        Assert.Equal(20D, positioned.X, 3);
        Assert.Equal(42D, positioned.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlAbsolutePosition_UsesInlineStaticAnchor() {
        const string html = "<p style='margin:0'>Before<span id='inline-auto' style='position:absolute;background:#ff0000'>Auto</span>After</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText automatic = FindText(rendered, "Auto");
        HtmlRenderText after = FindText(rendered, "After");
        Assert.Equal(after.X, automatic.X, 3);
        Assert.Equal(after.Y, automatic.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlAbsolutePosition_ResolvesInsetsAgainstWrappedInlineContainingBounds() {
        const string html = "<p style='width:70px;margin:0'><span style='position:relative'>Alpha Beta Gamma"
            + "<span id='inline-below' style='position:absolute;left:0;top:100%;width:10px;height:10px;background:#ff0000'></span></span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape positioned = FindPositionedShape(rendered, "span#inline-below");
        IReadOnlyList<HtmlRenderText> sourceText = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();
        Assert.NotEmpty(sourceText);
        Assert.Equal(sourceText.Min(text => text.X), positioned.X, 3);
        Assert.Equal(sourceText.Max(text => text.Y + text.Height), positioned.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlAbsolutePosition_UsesFlexAndGridStaticAlignmentWhenInsetsAreAuto() {
        const string html = "<div style='display:flex;width:200px;height:100px;justify-content:center;align-items:flex-end;margin:0'>"
            + "<div id='flex-auto' style='position:absolute;width:20px;height:20px;background:#ff0000'></div></div>"
            + "<div style='display:grid;width:200px;height:100px;justify-items:end;align-items:center;margin:0'>"
            + "<div id='grid-auto' style='position:absolute;width:20px;height:20px;background:#0000ff'></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 220D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape flex = FindPositionedShape(rendered, "div#flex-auto");
        HtmlRenderShape grid = FindPositionedShape(rendered, "div#grid-auto");
        Assert.Equal(90D, flex.X, 3);
        Assert.Equal(80D, flex.Y, 3);
        Assert.Equal(180D, grid.X, 3);
        Assert.Equal(140D, grid.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlFixedPosition_UsesInitialStaticAnchorOnEveryPage() {
        const string html = "<div style='height:30px;margin:0'>Before</div>"
            + "<div id='fixed-auto' style='position:fixed;width:20px;height:20px;margin:0;background:#ff0000'></div>"
            + "<div style='height:70px;margin:0'>PageOne</div><div style='height:70px;margin:0'>PageTwo</div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(120D / HtmlRenderOptions.CssPixelsPerInch, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.True(rendered.Pages.Count >= 2);
        foreach (HtmlRenderPage page in rendered.Pages) {
            HtmlRenderShape fixedShape = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#fixed-auto" && shape.Shape.FillColor.HasValue);
            Assert.Equal(10D, fixedShape.X, 3);
            Assert.Equal(40D, fixedShape.Y, 3);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlAbsolutePosition_AccumulatesStaticPositionThroughFlexAndGridItems() {
        const string html = "<div style='position:relative;width:200px;height:50px;margin:0'>"
            + "<div style='display:flex;width:200px;height:50px'>"
            + "<div style='flex:0 0 100px;box-sizing:border-box'></div>"
            + "<div style='flex:0 0 100px;box-sizing:border-box;padding:5px'><div style='height:10px'></div>"
            + "<div id='flex-nested-auto' style='position:absolute;width:10px;height:10px;background:#ff0000'></div></div></div></div>"
            + "<div style='position:relative;width:200px;height:50px;margin:0'>"
            + "<div style='display:grid;grid-template-columns:100px 100px;width:200px;height:50px'>"
            + "<div style='grid-column:2;padding:5px'><div style='height:10px'></div>"
            + "<div id='grid-nested-auto' style='position:absolute;width:10px;height:10px;background:#0000ff'></div></div></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 220D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape flex = FindPositionedShape(rendered, "div#flex-nested-auto");
        HtmlRenderShape grid = FindPositionedShape(rendered, "div#grid-nested-auto");
        Assert.Equal(105D, flex.X, 3);
        Assert.Equal(15D, flex.Y, 3);
        Assert.Equal(105D, grid.X, 3);
        Assert.Equal(65D, grid.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlAbsolutePosition_UsesDeclaredGridAreaAsContainingBlock() {
        const string html = "<div style='display:grid;grid-template-columns:100px 100px;grid-template-rows:50px 50px;width:200px;height:100px;margin:0'>"
            + "<div id='grid-area-auto' style='position:absolute;grid-column:2;grid-row:2;justify-self:end;align-self:center;width:20px;height:20px;background:#ff0000'></div>"
            + "<div id='grid-area-inset' style='position:absolute;grid-column:2;grid-row:2;left:10%;top:10%;width:20px;height:20px;background:#0000ff'></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 220D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape automatic = FindPositionedShape(rendered, "div#grid-area-auto");
        HtmlRenderShape inset = FindPositionedShape(rendered, "div#grid-area-inset");
        Assert.Equal(180D, automatic.X, 3);
        Assert.Equal(65D, automatic.Y, 3);
        Assert.Equal(110D, inset.X, 3);
        Assert.Equal(55D, inset.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlRelativePosition_StacksWithRootAbsoluteAndNormalFlowByNumericZIndex() {
        const string html = "<div id='root-normal' style='width:40px;height:40px;margin:0;background:#00ff00'></div>"
            + "<div id='root-relative-negative' style='position:relative;z-index:-1;top:-40px;width:40px;height:40px;margin:0;background:#ff0000'></div>"
            + "<div id='root-absolute-one' style='position:absolute;z-index:1;left:0;top:0;width:40px;height:40px;background:#0000ff'></div>"
            + "<div id='root-relative-two' style='position:relative;z-index:2;top:-80px;width:40px;height:40px;margin:0;background:#ffff00'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] paintSources = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("div#root-", StringComparison.Ordinal))
            .Select(shape => shape.Source!)
            .ToArray();
        Assert.Equal(new[] { "div#root-relative-negative", "div#root-normal", "div#root-absolute-one", "div#root-relative-two" }, paintSources);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Yellow, raster.GetPixel(20, 20));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionZIndexPending);
    }

    [Fact]
    public void HtmlRelativePosition_StacksInsideBlockFlexAndGridContainers() {
        const string html = "<div id='block-host' style='width:40px;height:80px;margin:0'>"
            + "<div id='block-normal' style='width:40px;height:40px;margin:0;background:#00ff00'></div>"
            + "<div id='block-negative' style='position:relative;z-index:-1;top:-40px;width:40px;height:40px;margin:0;background:#ff0000'></div></div>"
            + "<div id='flex-host' style='display:flex;width:40px;height:40px;margin:0'>"
            + "<div id='flex-normal-z' style='flex:0 0 40px;height:40px;background:#00ff00'></div>"
            + "<div id='flex-negative-z' style='position:relative;z-index:-1;left:-40px;flex:0 0 40px;height:40px;background:#ff0000'></div></div>"
            + "<div id='grid-host' style='display:grid;grid-template-columns:40px;width:40px;height:40px;margin:0'>"
            + "<div id='grid-normal-z' style='grid-area:1/1;width:40px;height:40px;background:#00ff00'></div>"
            + "<div id='grid-negative-z' style='position:relative;z-index:-1;grid-area:1/1;width:40px;height:40px;background:#ff0000'></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] blockOrder = PaintSources(rendered, "div#block-");
        string[] flexOrder = PaintSources(rendered, "div#flex-");
        string[] gridOrder = PaintSources(rendered, "div#grid-");
        Assert.Equal(new[] { "div#block-negative", "div#block-normal" }, blockOrder);
        Assert.Equal(new[] { "div#flex-negative-z", "div#flex-normal-z" }, flexOrder);
        Assert.Equal(new[] { "div#grid-negative-z", "div#grid-normal-z" }, gridOrder);
    }

    [Fact]
    public void HtmlRelativePosition_KeepsNestedContextBelowItsParentStackingLevel() {
        const string html = "<div style='width:40px;height:80px;margin:0'>"
            + "<div id='relative-sibling-five' style='position:relative;z-index:5;width:40px;height:40px;background:#0000ff'></div>"
            + "<div id='relative-parent-ten' style='position:relative;z-index:10;top:-40px;width:40px;height:40px;background:#ff0000'>"
            + "<div id='relative-nested-negative' style='position:relative;z-index:-100;width:40px;height:40px;background:#00ff00'></div></div></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] paintSources = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "div#relative-sibling-five" || shape.Source == "div#relative-parent-ten" || shape.Source == "div#relative-nested-negative")
            .Select(shape => shape.Source!)
            .ToArray();
        Assert.Equal(new[] { "div#relative-sibling-five", "div#relative-parent-ten", "div#relative-nested-negative" }, paintSources);
    }

    [Fact]
    public void HtmlRelativePosition_StacksInlineAtomicContextsByZIndex() {
        const string html = "<p style='margin:0;line-height:40px'>"
            + "<span style='display:inline-flex;width:40px;height:40px'><span id='inline-normal-box' style='width:40px;height:40px;background:#00ff00'></span></span>"
            + "<span style='display:inline-flex;position:relative;z-index:-1;left:-40px;width:40px;height:40px'><span id='inline-negative-box' style='width:40px;height:40px;background:#ff0000'></span></span>"
            + "</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] paintSources = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "span#inline-normal-box" || shape.Source == "span#inline-negative-box")
            .Select(shape => shape.Source!)
            .ToArray();
        Assert.Equal(new[] { "span#inline-negative-box", "span#inline-normal-box" }, paintSources);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0x00, 0xFF, 0x00), raster.GetPixel(20, 20));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionZIndexPending);
    }

    [Fact]
    public void HtmlPositioning_KeepsInlineAbsoluteLayerInsideParentInlineContext() {
        const string html = "<p style='margin:0;line-height:40px'>"
            + "<span style='display:inline-flex;position:relative;z-index:5;width:40px;height:40px'><span id='inline-sibling-five' style='width:40px;height:40px;background:#0000ff'></span></span>"
            + "<span style='position:relative;z-index:10;left:-40px'>"
            + "<span style='display:inline-flex;width:40px;height:40px'><span id='inline-parent-ten' style='width:40px;height:40px;background:#ff0000'></span></span>"
            + "<span id='inline-nested-negative' style='position:absolute;z-index:-100;left:0;top:0;width:40px;height:40px;background:#00ff00'></span></span>"
            + "</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] paintSources = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "span#inline-sibling-five" || shape.Source == "span#inline-parent-ten" || shape.Source == "span#inline-nested-negative")
            .Select(shape => shape.Source!)
            .ToArray();
        Assert.Equal(new[] { "span#inline-sibling-five", "span#inline-nested-negative", "span#inline-parent-ten" }, paintSources);
        Assert.All(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().Where(shape => paintSources.Contains(shape.Source)),
            shape => Assert.True(shape.X >= 0D, shape.Source + ":" + shape.X.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Red, raster.GetPixel(20, 20));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback
            || diagnostic.Code == HtmlRenderDiagnosticCodes.PositionZIndexPending);
    }

    [Fact]
    public void HtmlFixedPosition_UsesInlineStaticMarkerAndRepeatsWithoutConsumingLineWidth() {
        const string baselineHtml = "<p style='margin:0'>BeforeAfter</p><div style='height:70px'></div><div style='height:70px'></div>";
        const string positionedHtml = "<p style='margin:0'>Before<span id='fixed-inline-auto' style='position:fixed;width:10px;height:10px;background:#ff0000'></span>After</p>"
            + "<div style='height:70px'></div><div style='height:70px'></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(200D / HtmlRenderOptions.CssPixelsPerInch, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument baseline = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(baselineHtml), options);
        HtmlRenderDocument positioned = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(positionedHtml), options);
        HtmlRenderText baselineText = Assert.Single(baseline.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "BeforeAfter");
        IReadOnlyList<HtmlRenderText> positionedText = positioned.Pages[0].Visuals.OfType<HtmlRenderText>()
            .Where(text => text.Text == "Before" || text.Text == "After")
            .ToList();
        Assert.Equal(2, positionedText.Count);
        Assert.Equal(baselineText.Y, positionedText[0].Y, 3);
        Assert.Equal(baselineText.Y, positionedText[1].Y, 3);
        Assert.True(positioned.Pages.Count >= 2);
        double expectedX = positionedText.Single(text => text.Text == "After").X;
        foreach (HtmlRenderPage page in positioned.Pages) {
            HtmlRenderShape fixedShape = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "span#fixed-inline-auto" && shape.Shape.FillColor.HasValue);
            Assert.Equal(expectedX, fixedShape.X, 3);
            Assert.Equal(baselineText.Y, fixedShape.Y, 3);
        }
        Assert.DoesNotContain(positioned.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    [Fact]
    public void HtmlRelativePosition_NegativeRootStackingStillCountsAsPagedFlow() {
        const string html = "<div id='negative-page-one' style='position:relative;z-index:-1;height:60px;margin:0;background:#ff0000'></div>"
            + "<div id='normal-page-two' style='height:60px;margin:0;background:#0000ff'></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#negative-page-one");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#normal-page-two");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#normal-page-two");
    }

    [Fact]
    public void HtmlRelativePosition_ResolvesTrailingInsetsAndExplicitVerticalPercentages() {
        const string baselineHtml = "<div style='height:100px;margin:0'><span>Marker</span></div>";
        const string positionedHtml = "<div style='height:100px;margin:0'><span style='position:relative;right:5px;bottom:10%'>Marker</span></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument baseline = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(baselineHtml), options);
        HtmlRenderDocument positioned = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(positionedHtml), options);

        AssertPaintOffset(baseline, positioned, "Marker", -5D, -10D);
        Assert.DoesNotContain(positioned.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionInsetUnsupported);
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

    private static string[] PaintSources(HtmlRenderDocument document, string sourcePrefix) =>
        document.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith(sourcePrefix, StringComparison.Ordinal))
            .Select(shape => shape.Source!)
            .ToArray();

    private static (int PageNumber, HtmlRenderText Text) FindTextWithPage(HtmlRenderDocument document, string marker) {
        foreach (HtmlRenderPage page in document.Pages) {
            HtmlRenderText? text = page.Visuals.OfType<HtmlRenderText>().SingleOrDefault(item => item.Text == marker);
            if (text != null) return (page.PageNumber, text);
        }

        throw new Xunit.Sdk.XunitException("Rendered text marker was not found: " + marker);
    }
}
