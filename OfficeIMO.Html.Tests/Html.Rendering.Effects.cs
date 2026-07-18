using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTransform_TranslatesBlockPaintWithoutChangingFlow() {
        const string html = "<div id='translated' style='width:20px;height:20px;margin:0;background:#ff0000;transform-origin:0 0;transform:translate(30px,10px)'></div>"
            + "<div id='following' style='width:20px;height:20px;margin:0;background:#0000ff'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderEffectGroup group = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderEffectGroup>(), item => item.Source == "div#translated");
        HtmlRenderShape following = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), item => item.Source == "div#following" && item.Shape.FillColor.HasValue);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(30D, group.Transform.OffsetX, 3);
        Assert.Equal(10D, group.Transform.OffsetY, 3);
        Assert.Equal(20D, following.Y, 3);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(5, 5));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(35, 15));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(5, 25));
    }

    [Fact]
    public void HtmlTransform_ComposesFunctionsAndOriginInCssOrder() {
        const string html = "<div id='composed' style='width:10px;height:10px;margin:0;background:#00ff00;transform-origin:0 0;transform:translate(10px,5px) scale(2,1.5)'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        HtmlRenderEffectGroup group = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderEffectGroup>(), item => item.Source == "div#composed");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(2D, group.Transform.M11, 3);
        Assert.Equal(1.5D, group.Transform.M22, 3);
        Assert.Equal(10D, group.Transform.OffsetX, 3);
        Assert.Equal(5D, group.Transform.OffsetY, 3);
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(11, 6));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(28, 18));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(31, 18));
    }

    [Fact]
    public void HtmlOpacity_CompositesDescendantsAsOneIsolatedGroup() {
        const string html = "<div id='opacity-group' style='position:relative;width:20px;height:20px;margin:0;opacity:.5'>"
            + "<div style='position:absolute;left:0;top:0;width:20px;height:20px;background:#ff0000'></div>"
            + "<div style='position:absolute;left:0;top:0;width:20px;height:20px;background:#ff0000'></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        HtmlRenderEffectGroup group = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderEffectGroup>(), item => item.Source == "div#opacity-group");
        OfficeDrawingEffectGroup drawingGroup = Assert.Single(rendered.Pages[0].CreateDrawing().Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(10, 10);

        Assert.Equal(0.5D, group.Opacity, 3);
        Assert.Equal(0.5D, drawingGroup.Opacity, 3);
        Assert.Equal((byte)255, pixel.R);
        Assert.InRange(pixel.A, (byte)127, (byte)128);
    }

    [Fact]
    public void HtmlOpacity_AppliesOnceToGradientPaint() {
        const string html = "<div style='width:20px;height:20px;margin:0;background:linear-gradient(#ff0000,#ff0000);opacity:.5'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(10, 10);

        Assert.Equal((byte)255, pixel.R);
        Assert.InRange(pixel.A, (byte)127, (byte)128);
    }

    [Fact]
    public void HtmlEffects_FlowThroughPngSvgAndSearchablePdfWithLinks() {
        const string link = "https://example.com/effect";
        const string html = "<div style='width:90px;height:20px;margin:0;background:#ff0000;font-size:10px;line-height:10px;transform-origin:0 0;transform:translate(20px,5px);opacity:.75'><a href='https://example.com/effect'>EffectPdfMarker</a></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 140D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, options);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(140D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        PdfCore.PdfLogicalLinkAnnotation pdfLink = Assert.Single(PdfCore.PdfLogicalDocument.Load(pdf).GetLinksByUri(link));

        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(5, 5));
        Assert.True(raster.GetPixel(105, 10).A > 0);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("opacity=\"0.75\"", svg, StringComparison.Ordinal);
        Assert.Contains("matrix(1 0 0 1 20 5)", svg, StringComparison.Ordinal);
        Assert.Contains("EffectPdfMarker", pdfText, StringComparison.Ordinal);
        Assert.Contains("/Group << /S /Transparency /I true /K false >>", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.True(pdfLink.SourceLink.X1 >= 15D - 0.01D);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlEffects_NestedOpacityGroupsStayIsolatedAndSearchableInPdf() {
        const string html = "<div id='outer' style='position:relative;width:45px;height:35px;margin:0;opacity:.5;transform-origin:0 0;transform:translateX(5px)'>"
            + "<div id='inner' style='position:absolute;left:0;top:0;width:20px;height:20px;opacity:.5'>"
            + "<div style='position:absolute;left:0;top:0;width:20px;height:20px;background:#ff0000'></div>"
            + "<div style='position:absolute;left:0;top:0;width:20px;height:20px;background:#ff0000'></div>"
            + "</div><div style='position:absolute;left:0;top:24px;font-size:6px;line-height:7px'>NestedEffectMarker</div></div>";
        var renderOptions = new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), renderOptions);
        HtmlRenderEffectGroup outer = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderEffectGroup>(), item => item.Source == "div#outer");
        HtmlRenderEffectGroup inner = Assert.Single(EnumerateRenderVisuals(outer.Visuals).OfType<HtmlRenderEffectGroup>(), item => item.Source == "div#inner");
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(10, 10);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(renderOptions);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string rawPdf = Encoding.ASCII.GetString(pdf);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(0.5D, outer.Opacity, 3);
        Assert.Equal(0.5D, inner.Opacity, 3);
        Assert.Equal((byte)255, pixel.R);
        Assert.InRange(pixel.A, (byte)63, (byte)65);
        Assert.Contains("NestedEffectMarker", pdfText, StringComparison.Ordinal);
        Assert.True(rawPdf.Split(new[] { "/Group << /S /Transparency /I true /K false >>" }, StringSplitOptions.None).Length - 1 >= 2);
        Assert.DoesNotContain("OIMO_EFFECT_GROUP", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlEffects_RebaseAcrossPagedFragmentsWithoutDroppingPaint() {
        const string html = "<div id='paged-effect' style='width:20px;height:50px;margin:0;background:#0000ff;transform-origin:0 0;transform:translateX(10px);opacity:.75'></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            HtmlRenderEffectGroup group = Assert.Single(page.Visuals.OfType<HtmlRenderEffectGroup>(), item => item.Source == "div#paged-effect");
            OfficeColor pixel = OfficeDrawingRasterRenderer.Render(page.CreateDrawing()).GetPixel(15, 10);
            Assert.Equal(10D, group.Transform.OffsetX, 3);
            Assert.Equal((byte)255, pixel.B);
            Assert.InRange(pixel.A, (byte)190, (byte)192);
        });
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlEffects_InlineBlockUsesAnAtomicEffectGroup() {
        const string html = "<div style='width:60px;margin:0;font-size:10px;line-height:20px'>"
            + "<span id='atomic-effect' style='display:inline-block;width:10px;height:10px;margin:0;background:#ff0000;transform-origin:0 0;transform:translateX(10px);opacity:.5'></span>X</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 60D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        HtmlRenderEffectGroup group = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderEffectGroup>(), item => item.Source == "span#atomic-effect");
        HtmlRenderShape shape = Assert.Single(group.Visuals.OfType<HtmlRenderShape>(), item => item.Shape.FillColor.HasValue);
        OfficePoint sample = group.Transform.TransformPoint(new OfficePoint(shape.X + shape.Width / 2D, shape.Y + shape.Height / 2D));
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel((int)sample.X, (int)sample.Y);

        Assert.Equal(10D, group.Width, 3);
        Assert.Equal((byte)255, pixel.R);
        Assert.InRange(pixel.A, (byte)127, (byte)128);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.InlinePaintEffectUnsupported);
    }

    [Fact]
    public void HtmlEffects_InvalidAndInlineValuesUseCatalogedDiagnosticsAndSupportsTruth() {
        const string html = "<div id='invalid-effect' style='transform:warp(2);opacity:opaque'>Block</div>"
            + "<p><span id='inline-effect' style='transform:rotate(10deg);opacity:.5'>Inline</span></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported && item.Source == "div#invalid-effect");
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OpacityValueUnsupported && item.Source == "div#invalid-effect");
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.InlinePaintEffectUnsupported && item.Source == "span#inline-effect");
        Assert.Contains(HtmlRenderDiagnosticCodes.TransformValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.Contains(HtmlRenderDiagnosticCodes.OpacityValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.TransformValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(transform:translate(10px,20%))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(transform-origin:left top)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(opacity:50%)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(transform:warp(2))"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(transform-origin:left top 2px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(opacity:opaque)"));
    }
}
