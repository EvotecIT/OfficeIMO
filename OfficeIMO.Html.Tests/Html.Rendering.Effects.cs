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
    public void HtmlTransform_UsesTheActiveQueryContainerForLengthsAndOrigin() {
        const string html = "<section style='width:200px;height:100px;margin:0;container-type:size'>"
            + "<div id='translated' style='width:20px;height:20px;margin:0;background:red;transform-origin:0 0;transform:translate(10cqw,10cqh)'></div>"
            + "<div id='scaled' style='width:20px;height:20px;margin:0;background:blue;transform-origin:10cqw 10cqh;transform:scale(2)'></div>"
            + "</section>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 1000D,
            ViewportHeight = 500D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        HtmlRenderEffectGroup translated = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderEffectGroup>(), item => item.Source == "div#translated");
        HtmlRenderEffectGroup scaled = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderEffectGroup>(), item => item.Source == "div#scaled");

        Assert.Equal(20D, translated.Transform.OffsetX, 3);
        Assert.Equal(10D, translated.Transform.OffsetY, 3);
        Assert.Equal(-20D, scaled.Transform.OffsetX, 3);
        Assert.Equal(-30D, scaled.Transform.OffsetY, 3);
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
        PdfCore.PdfLogicalLinkAnnotation pdfLink = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(link));

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
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlEffects_InvalidAndInlineValuesUseCatalogedDiagnosticsAndSupportsTruth() {
        const string html = "<div id='invalid-effect' style='transform:warp(2);opacity:opaque'>Block</div>"
            + "<p><span id='invalid-inline-effect' style='transform:warp(2);opacity:opaque;clip-path:url(#missing)'>Inline</span></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported && item.Source == "div#invalid-effect");
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OpacityValueUnsupported && item.Source == "div#invalid-effect");
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TransformValueUnsupported && item.Source == "span#invalid-inline-effect");
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OpacityValueUnsupported && item.Source == "span#invalid-inline-effect");
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ClipPathValueUnsupported && item.Source == "span#invalid-inline-effect");
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

    [Fact]
    public void HtmlEffects_NonAtomicInlineTransformOpacityAndClipPathShareArtifactPipeline() {
        const string html = "<p style='width:100px;margin:0;font-size:10px;line-height:20px'>Before "
            + "<span id='inline-effect' style='background:#ff0000;transform-origin:0 0;transform:translateX(4px);opacity:.5;clip-path:inset(0 2px)'>InlineEffectMarker</span> after</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        IReadOnlyList<HtmlRenderVisual> visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToList();
        HtmlRenderEffectGroup effect = Assert.Single(visuals.OfType<HtmlRenderEffectGroup>(), group => group.Source == "span#inline-effect");
        HtmlRenderPathClipGroup clip = Assert.Single(visuals.OfType<HtmlRenderPathClipGroup>(), group => group.Source == "span#inline-effect");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = OfficeDrawingSvgExporter.ToSvg(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));

        Assert.Equal(4D, effect.Transform.OffsetX, 3);
        Assert.Equal(.5D, effect.Opacity, 3);
        Assert.True(clip.Width > 0D);
        Assert.Contains("opacity=\"0.5\"", svg, StringComparison.Ordinal);
        Assert.Contains("InlineEffectMarker", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(raster.Width > 0 && raster.Height > 0);
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("opacity:.5")]
    [InlineData("transform:translateX(1px)")]
    [InlineData("clip-path:inset(0)")]
    public void HtmlEffects_NonAtomicInlineStackingPreservesAuthoredTextOrder(string effect) {
        string html = "<p style='margin:0'><span style='" + effect + "'>A</span>B</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options) {
            PdfOptions = new PdfCore.PdfOptions { TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers }
        });

        Assert.Equal("AB", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))));
        Assert.Equal("AB", string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character))));
        Assert.True(PdfCore.PdfReadDocument.Open(pdf).HasTaggedContent);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlClipPath_BasicShapesShareOneVectorSceneAcrossRasterAndSvg() {
        const string html = "<div id='polygon' style='width:20px;height:20px;margin:0;background:red;clip-path:polygon(0 0,100% 0,0 100%)'></div>"
            + "<div id='inset' style='width:20px;height:20px;margin:0;background:blue;clip-path:inset(2px 3px 4px 5px round 2px)'></div>"
            + "<div id='circle' style='width:20px;height:20px;margin:0;background:green;clip-path:circle(8px at 50% 50%)'></div>"
            + "<div id='ellipse' style='width:20px;height:20px;margin:0;background:purple;clip-path:ellipse(40% 25% at center)'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 90D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        IReadOnlyList<HtmlRenderPathClipGroup> clips = EnumerateRenderVisuals(rendered.Pages[0].Visuals)
            .OfType<HtmlRenderPathClipGroup>()
            .Where(group => group.Source != null && group.Source.StartsWith("div#", StringComparison.Ordinal))
            .ToList();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = OfficeDrawingSvgExporter.ToSvg(rendered.Pages[0].CreateDrawing());

        Assert.Equal(4, clips.Count);
        HtmlRenderPathClipGroup inset = Assert.Single(clips, group => group.Source == "div#inset");
        Assert.Equal(5D, inset.X, 3);
        Assert.Equal(22D, inset.Y, 3);
        Assert.Equal(12D, inset.Width, 3);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(3, 3));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(18, 18));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(8, 25));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(2, 25));
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("<path", svg, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ClipPathValueUnsupported);
    }

    [Fact]
    public void HtmlClipPath_AllowsSignedInsetsCoordinatesAndPositions() {
        const string html = "<div id='inset' style='width:20px;height:20px;margin:0;background:red;clip-path:inset(-5px)'></div>"
            + "<div id='polygon' style='width:20px;height:20px;margin:0;background:blue;clip-path:polygon(-10px 0,20px 0,20px 20px)'></div>"
            + "<div id='circle' style='width:20px;height:20px;margin:0;background:green;clip-path:circle(5px at -2px 10px)'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 70D,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        IReadOnlyList<HtmlRenderPathClipGroup> clips = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderPathClipGroup>()
            .ToList();

        HtmlRenderPathClipGroup inset = Assert.Single(clips, group => group.Source == "div#inset");
        HtmlRenderPathClipGroup polygon = Assert.Single(clips, group => group.Source == "div#polygon");
        HtmlRenderPathClipGroup circle = Assert.Single(clips, group => group.Source == "div#circle");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = OfficeDrawingSvgExporter.ToSvg(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));
        Assert.Equal((-5D, 30D), (inset.X, inset.Width));
        Assert.Equal((-10D, 30D), (polygon.X, polygon.Width));
        Assert.Equal((-7D, 10D), (circle.X, circle.Width));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(2, 2));
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.Equal(1, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:circle(-5px))"));
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlClipPath_SignedInsetCrossingFarPageEdgesConvertsToPdf() {
        const string html = "<x-box id='far-edge' style='display:block;width:20px;height:20px;margin:15px 0 0 15px;background:red;clip-path:inset(-5px)'></x-box>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderPathClipGroup clip = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderPathClipGroup>(),
            group => group.Source == "x-box#far-edge");
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));

        Assert.True(clip.ClipX >= 0D && clip.ClipY >= 0D);
        Assert.True(clip.ClipPath.Width > 20D && clip.ClipPath.Height > 20D);
        Assert.True(clip.ClipX + clip.ClipPath.Width > rendered.Pages[0].Width);
        Assert.True(clip.ClipY + clip.ClipPath.Height > rendered.Pages[0].Height);
        Assert.Equal(1, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlClipPath_GeometryBoxesResolveShapePercentagesAndOrigins() {
        const string shared = "width:40px;height:20px;padding:10px;border:2px solid black;margin:5px;background:red;";
        string html = "<div id='content' style='" + shared + "clip-path:inset(0) content-box'></div>"
            + "<div id='padding' style='" + shared + "clip-path:padding-box inset(0)'></div>"
            + "<div id='border' style='" + shared + "clip-path:inset(0) border-box'></div>"
            + "<div id='margin' style='" + shared + "clip-path:margin-box'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        IReadOnlyList<HtmlRenderPathClipGroup> clips = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderPathClipGroup>()
            .ToList();

        HtmlRenderPathClipGroup content = Assert.Single(clips, group => group.Source == "div#content");
        HtmlRenderPathClipGroup padding = Assert.Single(clips, group => group.Source == "div#padding");
        HtmlRenderPathClipGroup border = Assert.Single(clips, group => group.Source == "div#border");
        HtmlRenderPathClipGroup margin = Assert.Single(clips, group => group.Source == "div#margin");
        Assert.Equal((17D, 40D, 20D), (content.X, content.Width, content.Height));
        Assert.Equal((7D, 60D, 40D), (padding.X, padding.Width, padding.Height));
        Assert.Equal((5D, 64D, 44D), (border.X, border.Width, border.Height));
        Assert.Equal((0D, 74D, 54D), (margin.X, margin.Width, margin.Height));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:circle(40%) content-box)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:padding-box inset(8px))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:margin-box)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:content-box inset(0) padding-box)"));
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlClipPath_EmptyBasicShapesRemainValidEmptyClipsAcrossArtifacts() {
        const string html = "<div id='inset-empty' style='width:20px;height:20px;margin:0;background:red;clip-path:inset(60%);font-size:8px'>InsetMarker</div>"
            + "<div id='circle-empty' style='width:20px;height:20px;margin:0;background:blue;clip-path:circle(0)'>CircleMarker</div>"
            + "<div id='ellipse-empty' style='width:20px;height:20px;margin:0;background:green;clip-path:ellipse(0 10px)'>EllipseMarker</div>"
            + "<div id='polygon-empty' style='width:20px;height:20px;margin:0;background:purple;clip-path:polygon(0 0,0 0,0 0)'>PolygonMarker</div>"
            + "<div id='polygon-collinear-empty' style='width:20px;height:20px;margin:0;background:orange;clip-path:polygon(0 0,10px 10px,20px 20px)'>CollinearMarker</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 110D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        IReadOnlyList<HtmlRenderPathClipGroup> clips = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderPathClipGroup>()
            .Where(group => group.Source != null && group.Source.Contains("empty", StringComparison.Ordinal))
            .ToList();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = OfficeDrawingSvgExporter.ToSvg(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(5, clips.Count);
        Assert.All(clips, clip => Assert.Equal(OfficeClipPathKind.Empty, clip.ClipPath.Kind));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(10, 10));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(10, 30));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(10, 50));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(10, 70));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(10, 90));
        Assert.Contains("<path d=\"\"", svg, StringComparison.Ordinal);
        Assert.Contains("InsetMarker", pdfText, StringComparison.Ordinal);
        Assert.Contains("CircleMarker", pdfText, StringComparison.Ordinal);
        Assert.Contains("EllipseMarker", pdfText, StringComparison.Ordinal);
        Assert.Contains("PolygonMarker", pdfText, StringComparison.Ordinal);
        Assert.Contains("CollinearMarker", pdfText, StringComparison.Ordinal);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:polygon(0 0,0 0,0 0))"));
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_NonRectangularClipSuppressesOnlyLinksOutsideTheVisibleRegion() {
        const string html = "<div style='width:100px;height:20px;margin:0;clip-path:polygon(0 0,100% 0,0 100%)'>"
            + "<a style='font-size:8px;line-height:10px' href='https://example.com/inside'>Inside</a></div>"
            + "<div style='width:100px;height:20px;margin:0;clip-path:polygon(100% 0,100% 100%,0 100%)'>"
            + "<a style='font-size:8px;line-height:10px' href='https://example.com/outside'>Outside</a></div>";
        var options = new HtmlPdfSaveOptions {
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(options);
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(pdf);

        Assert.Single(logical.GetLinksByUri("https://example.com/inside"));
        Assert.Empty(logical.GetLinksByUri("https://example.com/outside"));
    }

    [Fact]
    public void HtmlClipPath_DefaultCircleAndEllipseArgumentsUseCssDefaults() {
        const string html = "<div id='circle-default' style='width:20px;height:20px;margin:0;background:red;clip-path:circle()'></div>"
            + "<div id='ellipse-default' style='width:20px;height:20px;margin:0;background:blue;clip-path:ellipse()'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        IReadOnlyList<HtmlRenderPathClipGroup> clips = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderPathClipGroup>()
            .Where(group => group.Source != null && group.Source.EndsWith("-default", StringComparison.Ordinal))
            .ToList();

        Assert.Equal(2, clips.Count);
        Assert.All(clips, clip => Assert.Equal(OfficeClipPathKind.Path, clip.ClipPath.Kind));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:circle())"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:ellipse())"));
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("(clip-path:inset(10cqw))")]
    [InlineData("(clip-path:circle(10cqh at 20cqi 30cqb))")]
    [InlineData("(clip-path:polygon(0 0,100cqmin 0,0 100cqmax))")]
    public void HtmlClipPath_SupportsAcceptsContainerRelativeUnits(string condition) {
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports(condition));
    }

    [Fact]
    public void HtmlClipPath_CircleAndEllipseImplementTheCompleteBasicShapeRadialGrammar() {
        const string html = "<div style='width:40px;height:20px;background:red;clip-path:circle(closest-corner at 25% 25%)'></div>"
            + "<div style='width:40px;height:20px;background:red;clip-path:circle(farthest-corner at 25% 25%)'></div>"
            + "<div style='width:40px;height:20px;background:red;clip-path:circle(40% at 25% 25%)'></div>"
            + "<div style='width:40px;height:20px;background:red;clip-path:ellipse(closest-side at 25% 25%)'></div>"
            + "<div style='width:40px;height:20px;background:red;clip-path:ellipse(farthest-corner at 25% 25%)'></div>"
            + "<div style='width:40px;height:20px;background:red;clip-path:circle(closest-side at -5px 10px)'></div>"
            + "<div style='width:40px;height:20px;background:red;clip-path:ellipse(closest-side at -5px 10px)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 50D,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });

        IReadOnlyList<HtmlRenderPathClipGroup> clips = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderPathClipGroup>().ToList();
        Assert.Equal(7, clips.Count);
        Assert.All(clips, clip => Assert.Equal(OfficeClipPathKind.Path, clip.ClipPath.Kind));
        double expectedPercentageDiameter = 2D * 0.4D * Math.Sqrt((40D * 40D + 20D * 20D) / 2D);
        Assert.Equal(expectedPercentageDiameter, clips[2].ClipPath.Width, 3);
        Assert.Equal(expectedPercentageDiameter, clips[2].ClipPath.Height, 3);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:circle(closest-corner))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:circle(farthest-corner at 20px 10px))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:ellipse(closest-side))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:ellipse(farthest-corner at center))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:circle(40%))"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:circle(at))"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:ellipse(at))"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:ellipse(10px))"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:ellipse(closest-side farthest-side))"));
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlClipPath_PreservesSearchablePdfAndTruthfullyRejectsUnsupportedGeometry() {
        const string supported = "<div style='width:100px;height:30px;margin:0;background:#eee;clip-path:ellipse(50% 45% at center);font-size:10px'>ClipPathPdfMarker</div>";
        var options = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(120D / HtmlRenderOptions.CssPixelsPerInch, 50D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        };

        byte[] pdf = HtmlConversionDocument.Parse(supported).ToPdf(options);
        string extracted = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        Assert.Contains("ClipPathPdfMarker", extracted, StringComparison.Ordinal);

        HtmlRenderDocument fallback = HtmlRenderTestDriver.Render("<div style='clip-path:url(#missing)'>Fallback</div>");
        HtmlDiagnostic diagnostic = Assert.Single(fallback.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ClipPathValueUnsupported);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.True(fallback.HasLoss);
        Assert.Throws<HtmlConversionException>(() => HtmlRenderTestDriver.Render(
            "<div style='clip-path:path(\"M 0 0 L 1 1\")'>Fallback</div>",
            new HtmlRenderOptions { FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss }));
        Assert.Contains(HtmlRenderDiagnosticCodes.ClipPathValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.ClipPathValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:polygon(0 0,100% 0,0 100%))"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clip-path:url(#shape))"));
    }
}
