using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlImages_SvgPartialIntrinsicDimensionsUseViewBoxRatioInSharedLayout() {
        const string svgSource = "<svg xmlns='http://www.w3.org/2000/svg' width='200' viewBox='0 0 100 50'><rect width='100' height='50' fill='red'/></svg>";
        string data = Convert.ToBase64String(Encoding.UTF8.GetBytes(svgSource));
        string html = "<body style='margin:0'><img id='svg-image' src='data:image/svg+xml;base64," + data + "' alt='vector'></body>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 220D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(), item => item.Source == "img#svg-image");
        string exportedSvg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        Assert.Equal(200D, image.Width, 3);
        Assert.Equal(100D, image.Height, 3);
        Assert.Equal("image/svg+xml", image.ContentType);
        Assert.Contains("<image x=\"0\" y=\"0\" width=\"200\" height=\"100\"", exportedSvg, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlImages_ContainAndPositionFlowThroughPngSvgAndSearchablePdf() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<img id='contained' src='data:image/png;base64,{data}' alt='contained image' style='display:block;width:40px;height:40px;object-fit:contain;object-position:right bottom'>"
            + "<div style='font-size:6px;line-height:8px'>ImagePdf</div>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 60D,
            ViewportHeight = 55D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(60D / HtmlRenderOptions.CssPixelsPerInch, 55D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(0D, image.X, 3);
        Assert.Equal(20D, image.Y, 3);
        Assert.Equal(40D, image.Width, 3);
        Assert.Equal(20D, image.Height, 3);
        Assert.False(image.SourceCrop.HasCrop);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(20, 10));
        Assert.True(raster.GetPixel(20, 30).A > 0);
        Assert.Contains("<image x=\"0\" y=\"20\" width=\"40\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("ImagePdf", pdfText, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), extracted => extracted.IsImageFile && extracted.MimeType == "image/png");
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlImages_CoverUsesPositionedSourceCropAcrossSharedScene() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<img id='covered' src='data:image/png;base64,{data}' style='display:block;width:40px;height:40px;object-fit:cover;object-position:right center'>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 50D,
            ViewportHeight = 45D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderImage image = Assert.Single(HtmlRenderEngine.Render(html, options).Pages[0].Visuals.OfType<HtmlRenderImage>());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        Assert.Equal(0D, image.X, 3);
        Assert.Equal(0D, image.Y, 3);
        Assert.Equal(40D, image.Width, 3);
        Assert.Equal(40D, image.Height, 3);
        Assert.Equal(0.5D, image.SourceCrop.Left, 3);
        Assert.Equal(0D, image.SourceCrop.Top, 3);
        Assert.Equal(0D, image.SourceCrop.Right, 3);
        Assert.Equal(0D, image.SourceCrop.Bottom, 3);
        Assert.Contains("clipPath", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlImages_NoneScaleDownAspectRatioAndConstraintsUseIntrinsicGeometry() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<img id='none-fit' src='data:image/png;base64,{data}' style='display:block;width:40px;height:40px;object-fit:none;object-position:right bottom'>"
            + $"<img id='scaled-down' src='data:image/png;base64,{data}' style='display:block;width:10px;height:10px;object-fit:scale-down'>"
            + $"<img id='square-ratio' src='data:image/png;base64,{data}' style='display:block;width:30px;aspect-ratio:1/1;object-fit:fill'>"
            + $"<img id='max-size' src='data:image/png;base64,{data}' style='display:block;width:100px;max-width:40px'>"
            + $"<img id='edge-offset' src='data:image/png;base64,{data}' style='display:block;width:40px;height:40px;object-fit:none;object-position:right 5px bottom 4px'>"
            + $"<img id='border-box-image' src='data:image/png;base64,{data}' style='display:block;box-sizing:border-box;width:40px;height:40px;padding:5px;border:2px solid black'>";

        List<HtmlRenderImage> images = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 190D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        }).Pages[0].Visuals.OfType<HtmlRenderImage>().ToList();
        HtmlRenderImage none = Assert.Single(images, image => image.Source == "img#none-fit");
        HtmlRenderImage scaleDown = Assert.Single(images, image => image.Source == "img#scaled-down");
        HtmlRenderImage square = Assert.Single(images, image => image.Source == "img#square-ratio");
        HtmlRenderImage constrained = Assert.Single(images, image => image.Source == "img#max-size");
        HtmlRenderImage edgeOffset = Assert.Single(images, image => image.Source == "img#edge-offset");
        HtmlRenderImage borderBox = Assert.Single(images, image => image.Source == "img#border-box-image");

        Assert.Equal((20D, 30D, 20D, 10D), (none.X, none.Y, none.Width, none.Height));
        Assert.Equal((0D, 42.5D, 10D, 5D), (scaleDown.X, scaleDown.Y, scaleDown.Width, scaleDown.Height));
        Assert.Equal((0D, 50D, 30D, 30D), (square.X, square.Y, square.Width, square.Height));
        Assert.Equal((0D, 80D, 40D, 20D), (constrained.X, constrained.Y, constrained.Width, constrained.Height));
        Assert.Equal((15D, 126D, 20D, 10D), (edgeOffset.X, edgeOffset.Y, edgeOffset.Width, edgeOffset.Height));
        Assert.Equal((7D, 147D, 26D, 26D), (borderBox.X, borderBox.Y, borderBox.Width, borderBox.Height));
    }

    [Fact]
    public void HtmlImages_IntrinsicSizingFeedsFlexAndFloatPlanning() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<div style='display:flex;width:100px'><img id='flex-image' src='data:image/png;base64,{data}'><div id='flex-after' style='width:20px;height:10px;background:#0000ff'></div></div>"
            + $"<div style='display:flex;width:100px'><img id='flex-constrained' src='data:image/png;base64,{data}' style='width:100px;max-width:40px'><div id='flex-constrained-after' style='width:20px;height:10px;background:#00ff00'></div></div>"
            + $"<div style='width:100px;font-size:8px;line-height:10px'><img id='float-image' src='data:image/png;base64,{data}' style='float:left'><span>FloatText</span></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderImage flexImage = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(), image => image.Source == "img#flex-image");
        HtmlRenderShape flexAfter = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#flex-after");
        HtmlRenderImage flexConstrained = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(), image => image.Source == "img#flex-constrained");
        HtmlRenderShape flexConstrainedAfter = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#flex-constrained-after");
        HtmlRenderImage floatImage = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(), image => image.Source == "img#float-image");
        HtmlRenderText floatText = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FloatText");

        Assert.Equal(20D, flexImage.Width, 3);
        Assert.Equal(20D, flexAfter.X, 3);
        Assert.Equal(40D, flexConstrained.Width, 3);
        Assert.Equal(40D, flexConstrainedAfter.X, 3);
        Assert.Equal(20D, floatImage.Width, 3);
        Assert.Equal(20D, floatText.X, 3);
    }

    [Fact]
    public void HtmlImages_NormalInlineBoxesWrapAndParticipateInBaselineAcrossBackends() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<p id='inline-line' style='width:60px;margin:0;font-size:10px;line-height:12px'>Before<a href='https://example.com/image'><img id='inline-image' src='data:image/png;base64,{data}' alt='inline image' style='width:18px;height:14px;margin:0 2px;border:1px solid #0000ff;border-radius:4px;object-fit:cover'></a>After</p>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 65D,
            ViewportHeight = 45D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        IReadOnlyList<HtmlRenderVisual> flattened = EnumerateRenderVisuals(rendered.Pages[0].Visuals).ToList();
        HtmlRenderPathClipGroup clip = Assert.Single(flattened.OfType<HtmlRenderPathClipGroup>(), group => group.Source == "img#inline-image:content-clip");
        HtmlRenderImage image = Assert.Single(flattened.OfType<HtmlRenderImage>(), visual => visual.Source == "img#inline-image");
        HtmlRenderText before = Assert.Single(flattened.OfType<HtmlRenderText>(), text => text.Text == "Before");
        HtmlRenderText after = Assert.Single(flattened.OfType<HtmlRenderText>(), text => text.Text == "After");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(65D / HtmlRenderOptions.CssPixelsPerInch, 45D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(18D, clip.Width, 3);
        Assert.Equal(14D, clip.Height, 3);
        Assert.Equal(18D, image.Width, 3);
        Assert.Equal(14D, image.Height, 3);
        Assert.True(image.SourceCrop.Left > 0D);
        Assert.True(image.SourceCrop.Right > 0D);
        Assert.Equal("https://example.com/image", image.LinkUri);
        Assert.True(image.X > before.X);
        Assert.True(image.Y < before.Y);
        Assert.True(after.Y > before.Y);
        Assert.True(raster.GetPixel((int)Math.Round(image.X + image.Width / 2D), (int)Math.Round(image.Y + image.Height / 2D)).A > 0);
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("<image", svg, StringComparison.Ordinal);
        Assert.Contains("Before", pdfText, StringComparison.Ordinal);
        Assert.Contains("After", pdfText, StringComparison.Ordinal);
        Assert.Single(PdfCore.PdfLogicalDocument.Load(pdf).GetLinksByUri("https://example.com/image"));
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), extracted => extracted.IsImageFile && extracted.MimeType == "image/png");
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlImages_InvalidValuesAndRoundedClipUseSharedPathAndCatalogedDiagnostics() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<img id='invalid-image' src='data:image/png;base64,{data}' style='display:block;width:30px;height:20px;object-fit:stretch;object-position:sideways;aspect-ratio:0/1'>"
            + $"<img id='rounded-image' src='data:image/png;base64,{data}' style='display:block;width:30px;height:20px;border-radius:8px 2px / 3px 6px'>";

        var options = new HtmlImageExportOptions {
            ViewportWidth = 50D,
            ViewportHeight = 45D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlDiagnostic replaced = Assert.Single(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ReplacedElementValueUnsupported);
        HtmlRenderPathClipGroup rounded = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderPathClipGroup>(), group => group.Source == "img#rounded-image:content-clip");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        Assert.Equal("img#invalid-image", replaced.Source);
        Assert.Contains("object-fit=stretch", replaced.Detail, StringComparison.Ordinal);
        Assert.Contains("object-position=sideways", replaced.Detail, StringComparison.Ordinal);
        Assert.Contains("aspect-ratio=0/1", replaced.Detail, StringComparison.Ordinal);
        Assert.Equal(OfficeClipPathKind.Path, rounded.ClipPath.Kind);
        Assert.True(rounded.ClipPath.Commands.Count >= 10);
        Assert.Single(rounded.Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(0, 20));
        Assert.True(raster.GetPixel(15, 30).A > 0);
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("<path", svg, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
        Assert.Contains(HtmlRenderDiagnosticCodes.ReplacedElementValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.ReplacedElementValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(object-fit:cover)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(object-fit:scale-down)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(object-position:right 4px bottom 2px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(aspect-ratio:auto 16 / 9)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(object-fit:stretch)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(object-position:left right)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(aspect-ratio:0/1)"));
    }

    [Fact]
    public void HtmlImages_RoundedRepeatedBackgroundUsesSharedPathClipAcrossPngSvgAndPdf() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 4));
        string html = $"<div id='rounded-background' style='width:30px;height:20px;border-radius:10px 2px / 4px 8px;background-image:url(data:image/png;base64,{data});background-size:4px 4px;background-repeat:repeat'></div>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 35D,
            ViewportHeight = 25D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderPathClipGroup group = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderPathClipGroup>(), item => item.Source == "div#rounded-background:background-image:clip");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(35D / HtmlRenderOptions.CssPixelsPerInch, 25D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);

        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.True(group.ClipPath.Commands.Count >= 10);
        Assert.Single(group.Visuals.OfType<HtmlRenderImagePattern>());
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(0, 0));
        Assert.True(raster.GetPixel(15, 10).A > 0);
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("<path", svg, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }
}
