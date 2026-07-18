using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlViewportOverflow_DocumentRootClipsSharedContinuousAndPdfContent() {
        const string html = "<style>html{overflow:hidden}body{overflow:visible}</style><div style='width:100px;height:12px;margin:0;font-size:6px;line-height:8px'>RootPdf</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderClipGroup viewport = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderClipGroup>(), group => group.Source == "html:viewport-overflow");
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(50D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.True(viewport.ClipHorizontal);
        Assert.True(viewport.ClipVertical);
        Assert.Equal(50D, viewport.ClipWidth, 3);
        Assert.Equal(30D, viewport.ClipHeight, 3);
        Assert.Contains("clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("RootPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.OverflowScrollSnapshot);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlViewportOverflow_UsesBodyWhenDocumentRootRemainsVisible() {
        const string html = "<style>html{overflow:visible}body{overflow-x:clip;overflow-y:visible}</style><div style='width:80px;height:10px;margin:0'>Body overflow</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 25D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        HtmlRenderClipGroup viewport = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderClipGroup>(), group => group.Source == "body:viewport-overflow");

        Assert.True(viewport.ClipHorizontal);
        Assert.False(viewport.ClipVertical);
        Assert.DoesNotContain(viewport.Visuals.OfType<HtmlRenderClipGroup>(), group => group.Source == "html:viewport-overflow");
    }

    [Fact]
    public void HtmlOverflowHidden_ClipsPositionedContentInSharedScene() {
        const string html = "<div id='clip' style='position:relative;width:40px;height:30px;margin:0;overflow:hidden'>"
            + "<div id='overflowing' style='position:absolute;left:20px;top:10px;width:40px;height:40px;background:#ff0000;color:#ff0000'>ClippedPdfMarker</div></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderClipGroup group = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(), item => item.Source == "div#clip");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(80D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.True(group.ClipHorizontal);
        Assert.True(group.ClipVertical);
        Assert.Equal(40D, group.ClipWidth, 3);
        Assert.Equal(30D, group.ClipHeight, 3);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 15));
        Assert.Equal(OfficeColor.White, raster.GetPixel(45, 15));
        Assert.Equal(OfficeColor.White, raster.GetPixel(25, 35));
        Assert.Contains("clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("ClippedPdfMarker", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))), StringComparison.Ordinal);
        Assert.Contains("ClippedPdfMarker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlOverflowVisible_PreservesPaintOutsideTheBox() {
        const string html = "<div style='position:relative;width:40px;height:30px;margin:0;overflow:visible'>"
            + "<div style='position:absolute;left:20px;top:10px;width:40px;height:40px;background:#ff0000'></div></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(45, 15));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 35));
        Assert.Empty(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>());
    }

    [Fact]
    public void HtmlOverflowHidden_PreservesClipOriginAwayFromPageOrigin() {
        const string html = "<div id='offset-clip' style='position:relative;width:40px;height:30px;margin:0 0 0 20px;overflow:hidden'>"
            + "<div style='position:absolute;left:20px;top:0;width:40px;height:30px;background:#ff0000'></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderClipGroup group = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(), item => item.Source == "div#offset-clip");
        OfficeDrawing drawing = rendered.Pages[0].CreateDrawing();
        OfficeDrawingGroup drawingGroup = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(20D, group.ClipX, 3);
        Assert.Equal(20D, drawingGroup.X, 3);
        Assert.Equal(-20D, drawingGroup.ContentOffsetX, 3);
        Assert.Equal(OfficeColor.White, raster.GetPixel(25, 10));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(45, 10));
        Assert.Equal(OfficeColor.White, raster.GetPixel(65, 10));
    }

    [Fact]
    public void HtmlOverflowClip_CanConstrainOneAxisOnly() {
        const string html = "<div id='clip-x' style='position:relative;width:40px;height:30px;margin:0;overflow-x:clip;overflow-y:visible'>"
            + "<div style='position:absolute;left:20px;top:20px;width:40px;height:30px;background:#ff0000'></div></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderClipGroup group = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(), item => item.Source == "div#clip-x");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.True(group.ClipHorizontal);
        Assert.False(group.ClipVertical);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(35, 40));
        Assert.Equal(OfficeColor.White, raster.GetPixel(45, 25));
    }

    [Theory]
    [InlineData("content-box", 11D, 28D)]
    [InlineData("padding-box", 8D, 34D)]
    [InlineData("border-box", 6D, 38D)]
    public void HtmlOverflowClipMargin_ResolvesVisualBoxEdges(string visualBox, double expectedX, double expectedWidth) {
        string html = "<body style='margin:0'><div id='clip-margin' style='position:relative;width:20px;height:20px;margin:10px;padding:3px;border:2px solid black;overflow:clip;overflow-clip-margin:"
            + visualBox + " 4px'>X</div></body>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 60D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderClipGroup group = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(),
            item => item.Source == "div#clip-margin");

        Assert.Equal(expectedX, group.ClipX, 3);
        Assert.Equal(expectedWidth, group.ClipWidth, 3);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OverflowClipMarginValueUnsupported);
    }

    [Fact]
    public void HtmlOverflowClipMargin_ExpandsOnlyClipOverflowAcrossBackends() {
        const string clipHtml = "<body style='margin:0'><div id='expanded' style='position:relative;width:20px;height:20px;margin:10px;overflow:clip;overflow-clip-margin:5px'>"
            + "<span style='position:absolute;left:-4px;top:0;width:3px;height:20px;background:red'></span></div></body>";
        const string hiddenHtml = "<body style='margin:0'><div id='hidden' style='width:20px;height:20px;margin:10px;overflow:hidden;overflow-clip-margin:5px'>X</div></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument clipped = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(clipHtml), options);
        HtmlRenderDocument hidden = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(hiddenHtml), options);
        HtmlRenderClipGroup clippedGroup = Assert.Single(EnumerateRenderVisuals(clipped.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(), item => item.Source == "div#expanded");
        HtmlRenderClipGroup hiddenGroup = Assert.Single(EnumerateRenderVisuals(hidden.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(), item => item.Source == "div#hidden");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(clipped.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(clipHtml).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(clipHtml).ToPdf(pdfOptions);

        Assert.Equal(5D, clippedGroup.ClipX, 3);
        Assert.Equal(30D, clippedGroup.ClipWidth, 3);
        Assert.Equal(10D, hiddenGroup.ClipX, 3);
        Assert.Equal(20D, hiddenGroup.ClipWidth, 3);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(7, 15));
        Assert.Contains("clipPath", svg, StringComparison.Ordinal);
        Assert.Equal(1, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(clipHtml).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlOverflowAuto_ReportsOnlyAnActualStaticScrollSnapshot() {
        const string overflowingHtml = "<div id='auto-overflow' style='position:relative;width:40px;height:30px;margin:0;overflow:auto'>"
            + "<div style='position:absolute;left:20px;top:10px;width:40px;height:40px;background:#ff0000'></div></div>";
        const string fittingHtml = "<div id='auto-fit' style='width:40px;height:30px;margin:0;overflow:auto'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument overflowing = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(overflowingHtml), options);
        HtmlRenderDocument fitting = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(fittingHtml), options);

        HtmlDiagnostic diagnostic = Assert.Single(overflowing.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OverflowScrollSnapshot);
        Assert.Equal("div#auto-overflow", diagnostic.Source);
        Assert.DoesNotContain(fitting.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OverflowScrollSnapshot);
    }

    [Fact]
    public void HtmlOverflowScroll_ReportsInitialSnapshotEvenWhenContentFits() {
        const string html = "<div id='scroll-box' style='width:40px;height:30px;margin:0;overflow:scroll'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Single(rendered.Diagnostics, item =>
            item.Code == HtmlRenderDiagnosticCodes.OverflowScrollSnapshot
            && item.Source == "div#scroll-box");
    }

    [Fact]
    public void HtmlOverflow_NestedClipsIntersectAcrossBackends() {
        const string html = "<div id='outer-clip' style='position:relative;width:40px;height:40px;margin:0;overflow:hidden'>"
            + "<div id='inner-clip' style='position:absolute;left:20px;top:20px;width:40px;height:40px;overflow:hidden'>"
            + "<div style='position:absolute;left:10px;top:10px;width:30px;height:30px;background:#0000ff'></div></div></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D)
        }).Bytes);

        Assert.Equal(2, EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>().Count());
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(35, 35));
        Assert.Equal(OfficeColor.White, raster.GetPixel(45, 35));
        Assert.True(svg.Split(new[] { "<clipPath" }, StringSplitOptions.None).Length >= 3);
    }

    [Fact]
    public void HtmlOverflowHidden_ClipsFlexAndGridContainers() {
        string[] documents = {
            "<div id='layout-clip' style='display:flex;width:40px;height:20px;margin:0;overflow:hidden'><div style='flex:0 0 60px;height:20px;background:#ff0000'></div></div>",
            "<div id='layout-clip' style='display:grid;grid-template-columns:60px;width:40px;height:20px;margin:0;overflow:hidden'><div style='height:20px;background:#ff0000'></div></div>"
        };
        foreach (string html in documents) {
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
                ViewportWidth = 80D,
                ViewportHeight = 40D,
                Margins = HtmlRenderMargins.All(0D)
            });
            HtmlRenderClipGroup group = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(), item => item.Source == "div#layout-clip");
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
            Assert.True(group.ClipHorizontal);
            Assert.True(group.ClipVertical);
            Assert.Equal(OfficeColor.Red, raster.GetPixel(35, 10));
            Assert.Equal(OfficeColor.White, raster.GetPixel(45, 10));
        }
    }

    [Fact]
    public void HtmlOverflowClipGroup_FragmentsWithoutDroppingVectorContent() {
        const string html = "<div id='paged-clip' style='width:20px;height:50px;margin:0;overflow:hidden'>"
            + "<div style='width:20px;height:50px;background:#0000ff'></div></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            Assert.Contains(EnumerateRenderVisuals(page.Visuals).OfType<HtmlRenderClipGroup>(), group => group.Source == "div#paged-clip");
            Assert.Equal(OfficeColor.Blue, OfficeDrawingRasterRenderer.Render(page.CreateDrawing()).GetPixel(10, 10));
        });
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlOverflow_InvalidValuesUseCatalogedDiagnosticsAndSupportsTruth() {
        const string html = "<div id='invalid-overflow' style='overflow:sideways banana'>Text</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OverflowValueUnsupported);
        Assert.Equal("div#invalid-overflow", diagnostic.Source);
        Assert.Contains("overflow-x=sideways", diagnostic.Detail);
        Assert.Contains("overflow-y=banana", diagnostic.Detail);
        Assert.Contains(HtmlRenderDiagnosticCodes.OverflowValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.Contains(HtmlRenderDiagnosticCodes.OverflowScrollSnapshot, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.OverflowValueUnsupported, out _));
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.OverflowScrollSnapshot, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(overflow:hidden auto)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(overflow:sideways)"));
    }

    [Fact]
    public void HtmlOverflowClipMargin_InvalidValuesUseInitialFallbackAndCatalogedDiagnostics() {
        const string html = "<div id='invalid-clip-margin' style='overflow:clip;overflow-clip-margin:border-box -2px'>Text</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OverflowClipMarginValueUnsupported);

        Assert.Equal("div#invalid-clip-margin", diagnostic.Source);
        Assert.Contains("overflow-clip-margin=border-box -2px", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains(HtmlRenderDiagnosticCodes.OverflowClipMarginValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.OverflowClipMarginValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(overflow-clip-margin:content-box 3px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(overflow-clip-margin:4px border-box)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(overflow-clip-margin:padding-box -1px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(overflow-clip-margin:10%)"));
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateRenderVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clipGroup
                ? clipGroup.Visuals
                : visual is HtmlRenderPathClipGroup pathClipGroup
                    ? pathClipGroup.Visuals
                    : visual is HtmlRenderEffectGroup effectGroup ? effectGroup.Visuals
                    : visual is HtmlRenderSemanticGroup semanticGroup ? semanticGroup.Visuals
                    : visual is HtmlRenderLogicalTextGroup logicalTextGroup ? logicalTextGroup.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateRenderVisuals(children)) yield return child;
        }
    }
}
