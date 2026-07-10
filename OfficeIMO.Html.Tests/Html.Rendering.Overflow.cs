using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlOverflowHidden_ClipsPositionedContentInSharedScene() {
        const string html = "<div id='clip' style='position:relative;width:40px;height:30px;margin:0;overflow:hidden'>"
            + "<div id='overflowing' style='position:absolute;left:20px;top:10px;width:40px;height:40px;background:#ff0000;color:#ff0000'>ClippedPdfMarker</div></div>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderClipGroup group = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(), item => item.Source == "div#clip");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(80D / HtmlRenderOptions.CssPixelsPerInch, 60D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(html.SaveAsPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

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
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
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

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(45, 15));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 35));
        Assert.Empty(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>());
    }

    [Fact]
    public void HtmlOverflowHidden_PreservesClipOriginAwayFromPageOrigin() {
        const string html = "<div id='offset-clip' style='position:relative;width:40px;height:30px;margin:0 0 0 20px;overflow:hidden'>"
            + "<div style='position:absolute;left:20px;top:0;width:40px;height:30px;background:#ff0000'></div></div>";
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderClipGroup group = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderClipGroup>(), item => item.Source == "div#clip-x");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.True(group.ClipHorizontal);
        Assert.False(group.ClipVertical);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(35, 40));
        Assert.Equal(OfficeColor.White, raster.GetPixel(45, 25));
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

        HtmlRenderDocument overflowing = HtmlRenderEngine.Render(overflowingHtml, options);
        HtmlRenderDocument fitting = HtmlRenderEngine.Render(fittingHtml, options);

        HtmlDiagnostic diagnostic = Assert.Single(overflowing.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OverflowScrollSnapshot);
        Assert.Equal("div#auto-overflow", diagnostic.Source);
        Assert.DoesNotContain(fitting.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OverflowScrollSnapshot);
    }

    [Fact]
    public void HtmlOverflowScroll_ReportsInitialSnapshotEvenWhenContentFits() {
        const string html = "<div id='scroll-box' style='width:40px;height:30px;margin:0;overflow:scroll'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Single(rendered.Diagnostics.Diagnostics, item =>
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

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, new HtmlImageExportOptions {
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
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            Assert.Contains(EnumerateRenderVisuals(page.Visuals).OfType<HtmlRenderClipGroup>(), group => group.Source == "div#paged-clip");
            Assert.Equal(OfficeColor.Blue, OfficeDrawingRasterRenderer.Render(page.CreateDrawing()).GetPixel(10, 10));
        });
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlOverflow_InvalidValuesUseCatalogedDiagnosticsAndSupportsTruth() {
        const string html = "<div id='invalid-overflow' style='overflow:sideways banana'>Text</div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OverflowValueUnsupported);
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

    private static IEnumerable<HtmlRenderVisual> EnumerateRenderVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clipGroup
                ? clipGroup.Visuals
                : visual is HtmlRenderEffectGroup effectGroup ? effectGroup.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateRenderVisuals(children)) yield return child;
        }
    }
}
