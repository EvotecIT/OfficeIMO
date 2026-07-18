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
        var options = new HtmlRenderOptions {
            ViewportWidth = 220D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderDrawing image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>(), item => item.Source == "img#svg-image");
        string exportedSvg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        Assert.Equal(200D, image.Width, 3);
        Assert.Equal(100D, image.Height, 3);
        Assert.Single(image.Drawing.Shapes);
        Assert.Contains("<rect", exportedSvg, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/svg+xml", exportedSvg, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported);
    }

    [Fact]
    public void HtmlImages_SvgPrimitivesAndLocalReferencesFlowAsNativeVectorsAcrossPngSvgAndSearchablePdf() {
        const string svgSource = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'><defs><symbol id='marker' viewBox='0 0 2 2'><rect width='2' height='2' fill='currentColor'/></symbol></defs><path d='M0 0h20v20H0z' fill='red'/><circle cx='30' cy='10' r='8' fill='blue'/><path d='M22 10A8 6 30 0 1 38 10' fill='none' stroke='black'/><use href='#marker' style='fill:red;color:lime' transform='translate(18 8) scale(2)'/><text x='20' y='18' font-size='4' text-anchor='middle' textLength='20' lengthAdjust='spacingAndGlyphs' fill='black' transform='translate(0 -1)'>Svg<tspan font-weight='bold'>LabelX</tspan></text></svg>";
        string data = Convert.ToBase64String(Encoding.UTF8.GetBytes(svgSource));
        string html = "<body style='margin:0'><img id='vector' src='data:image/svg+xml;base64," + data + "' style='display:block;width:80px;height:40px'><div style='font-size:6px;line-height:8px'>SvgPdf</div></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 90D,
            ViewportHeight = 55D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderDrawing vector = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>(), visual => visual.Source == "img#vector");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string exportedSvg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(90D / HtmlRenderOptions.CssPixelsPerInch, 55D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(3, vector.Drawing.Shapes.Count);
        string[] svgTextRuns = vector.Drawing.Elements.OfType<OfficeDrawingEffectGroup>()
            .SelectMany(group => group.Drawing.Elements.OfType<OfficeDrawingText>())
            .Select(text => text.Text)
            .ToArray();
        Assert.Equal(new[] { "Svg", "LabelX" }, svgTextRuns);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(10, 10));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(60, 20));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(40, 20));
        Assert.Contains("<path", exportedSvg, StringComparison.Ordinal);
        Assert.Contains("<ellipse", exportedSvg, StringComparison.Ordinal);
        Assert.Contains("transform=\"matrix(", exportedSvg, StringComparison.Ordinal);
        Assert.Contains(">Svg</text>", exportedSvg, StringComparison.Ordinal);
        Assert.Contains(">LabelX</text>", exportedSvg, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/svg+xml", exportedSvg, StringComparison.Ordinal);
        Assert.Contains("SvgPdf", pdfText, StringComparison.Ordinal);
        Assert.Contains("SvgLabelX", pdfText, StringComparison.Ordinal);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlImages_SvgPaintServersStayNativeAcrossPngSvgAndSearchablePdf() {
        const string svgSource = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'><defs>"
            + "<linearGradient id='linear' gradientUnits='userSpaceOnUse' spreadMethod='repeat' x1='0' y1='0' x2='12.5%' y2='0'><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient>"
            + "<radialGradient id='radial' gradientUnits='userSpaceOnUse' gradientTransform='matrix(.5 0 0 1 15 0)' cx='30' cy='10' r='8' fx='28' fy='10'><stop offset='0' stop-color='white'/><stop offset='1' stop-color='navy'/></radialGradient>"
            + "</defs><rect width='10' height='20' fill='url(#linear)'/><rect x='10' width='10' height='20' fill='url(#linear)'/><rect x='20' width='20' height='20' fill='url(#radial)'/></svg>";
        string data = Convert.ToBase64String(Encoding.UTF8.GetBytes(svgSource));
        string html = "<body style='margin:0'><img id='paint-server' src='data:image/svg+xml;base64," + data
            + "' style='display:block;width:80px;height:40px'><div style='font-size:6px;line-height:8px'>SvgPdfX</div></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 90D,
            ViewportHeight = 55D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderDrawing visual = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>(), item => item.Source == "img#paint-server");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string exportedSvg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(90D / HtmlRenderOptions.CssPixelsPerInch, 55D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        OfficeLinearGradient first = Assert.IsType<OfficeLinearGradient>(visual.Drawing.Shapes[0].Shape.FillGradient);
        OfficeLinearGradient second = Assert.IsType<OfficeLinearGradient>(visual.Drawing.Shapes[1].Shape.FillGradient);
        Assert.Equal(0D, first.StartX, 8);
        Assert.Equal(1D, first.EndX, 8);
        Assert.Equal(0D, second.StartX, 8);
        Assert.Equal(1D, second.EndX, 8);
        Assert.True(first.Stops.Count > 2);
        Assert.True(second.Stops.Count > 2);
        OfficeRadialGradient radial = Assert.IsType<OfficeRadialGradient>(visual.Drawing.Shapes[2].Shape.FillRadialGradient);
        Assert.Equal(0.2D, radial.EndRadiusX);
        Assert.Equal(0.4D, radial.EndRadiusY);
        OfficeColor repeatStart = raster.GetPixel(3, 20);
        OfficeColor repeatEnd = raster.GetPixel(37, 20);
        Assert.True(repeatStart.R > repeatStart.B, $"Expected the repeated gradient start to be red-dominant, got {repeatStart}.");
        Assert.True(repeatEnd.B > repeatEnd.R, $"Expected the repeated gradient end to be blue-dominant, got {repeatEnd}.");
        Assert.Contains("<linearGradient", exportedSvg, StringComparison.Ordinal);
        Assert.Contains("<radialGradient", exportedSvg, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/svg+xml", exportedSvg, StringComparison.Ordinal);
        Assert.Contains("SvgPdfX", pdfText, StringComparison.Ordinal);
        Assert.Contains("/Shading", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlImages_SvgUnsupportedFeaturesAreDiagnosedWhilePrimitivesRemain() {
        const string svgSource = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'><rect width='20' height='20' fill='lime'/><text x='1' y='10'>Pending</text></svg>";
        string data = Convert.ToBase64String(Encoding.UTF8.GetBytes(svgSource));
        string html = "<img id='partial-svg' src='data:image/svg+xml;base64," + data + "'>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported);
        Assert.Equal("img#partial-svg", diagnostic.Source);
        Assert.Contains("features=1", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains(HtmlRenderDiagnosticCodes.SvgContentUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.SvgContentUnsupported, out _));
    }

    [Fact]
    public void HtmlImages_ContainAndPositionFlowThroughPngSvgAndSearchablePdf() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<img id='contained' src='data:image/png;base64,{data}' alt='contained image' style='display:block;width:40px;height:40px;object-fit:contain;object-position:right bottom'>"
            + "<div style='font-size:6px;line-height:8px'>ImagePdf</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 60D,
            ViewportHeight = 55D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(60D / HtmlRenderOptions.CssPixelsPerInch, 55D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

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
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlImages_CoverUsesPositionedSourceCropAcrossSharedScene() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<img id='covered' src='data:image/png;base64,{data}' style='display:block;width:40px;height:40px;object-fit:cover;object-position:right center'>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 45D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderImage image = Assert.Single(HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options).Pages[0].Visuals.OfType<HtmlRenderImage>());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

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
    public void HtmlImages_SvgCoverPreservesPositionedSourceCropAcrossSharedScene() {
        const string svgSource = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10'><rect width='10' height='10' fill='red'/><rect x='10' width='10' height='10' fill='blue'/></svg>";
        string data = Convert.ToBase64String(Encoding.UTF8.GetBytes(svgSource));
        string html = "<img id='covered-svg' src='data:image/svg+xml;base64," + data + "' style='display:block;width:10px;height:10px;object-fit:cover;object-position:right center'>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 12D,
            ViewportHeight = 12D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<HtmlRenderVisual> visuals = EnumerateRenderVisuals(rendered.Pages[0].Visuals).ToList();
        HtmlRenderClipGroup clip = Assert.Single(visuals.OfType<HtmlRenderClipGroup>(), item => item.Source == "img#covered-svg:object-fit-clip");
        HtmlRenderDrawing drawing = Assert.Single(clip.Visuals.OfType<HtmlRenderDrawing>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(10D, clip.Width, 3);
        Assert.Equal(20D, drawing.Width, 3);
        Assert.Equal(-10D, drawing.X, 3);
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(5, 5));
    }

    [Theory]
    [InlineData("image/gif")]
    [InlineData("image/bmp")]
    public void HtmlPdf_DirectRenderer_ConvertsDependencyFreeRasterFormatsToPng(string contentType) {
        byte[] bytes = contentType == "image/gif"
            ? Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==")
            : CreateSinglePixelBmp(0x12, 0x34, 0x56);
        string html = "<img src='data:" + contentType + ";base64," + Convert.ToBase64String(bytes) + "' style='width:10px;height:10px'>";

        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());
        IReadOnlyList<PdfCore.PdfExtractedImage> images = PdfCore.PdfImageExtractor.ExtractImages(pdf);

        Assert.Contains(images, image => image.IsImageFile && image.MimeType == "image/png");
    }

    private static byte[] CreateSinglePixelBmp(byte red, byte green, byte blue) {
        var bytes = new byte[58];
        bytes[0] = (byte)'B';
        bytes[1] = (byte)'M';
        BitConverter.GetBytes(bytes.Length).CopyTo(bytes, 2);
        BitConverter.GetBytes(54).CopyTo(bytes, 10);
        BitConverter.GetBytes(40).CopyTo(bytes, 14);
        BitConverter.GetBytes(1).CopyTo(bytes, 18);
        BitConverter.GetBytes(1).CopyTo(bytes, 22);
        BitConverter.GetBytes((short)1).CopyTo(bytes, 26);
        BitConverter.GetBytes((short)24).CopyTo(bytes, 28);
        BitConverter.GetBytes(4).CopyTo(bytes, 34);
        bytes[54] = blue;
        bytes[55] = green;
        bytes[56] = red;
        return bytes;
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

        List<HtmlRenderImage> images = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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
        var options = new HtmlRenderOptions {
            ViewportWidth = 65D,
            ViewportHeight = 45D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<HtmlRenderVisual> flattened = EnumerateRenderVisuals(rendered.Pages[0].Visuals).ToList();
        HtmlRenderPathClipGroup clip = Assert.Single(flattened.OfType<HtmlRenderPathClipGroup>(), group => group.Source == "img#inline-image:content-clip");
        HtmlRenderImage image = Assert.Single(flattened.OfType<HtmlRenderImage>(), visual => visual.Source == "img#inline-image");
        HtmlRenderText before = Assert.Single(flattened.OfType<HtmlRenderText>(), text => text.Text == "Before");
        HtmlRenderText after = Assert.Single(flattened.OfType<HtmlRenderText>(), text => text.Text == "After");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(65D / HtmlRenderOptions.CssPixelsPerInch, 45D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

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
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlImages_OversizedRasterContinuesThroughClippedPageFragments() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0));
        string html = "<html style='margin:0'><body style='margin:0'><img id='tall' src='data:image/png;base64," + data
            + "' style='display:block;width:40px;height:90px'></body></html>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<OfficeImageExportResult> pngPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Png, options);
        IReadOnlyList<OfficeImageExportResult> svgPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Svg, options);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            HtmlRenderClipGroup fragment = Assert.Single(page.Visuals.OfType<HtmlRenderClipGroup>(), group => group.Source == "img#tall");
            Assert.Single(fragment.Visuals.OfType<HtmlRenderImage>(), image => image.Source == "img#tall");
        });
        Assert.Equal(3, pngPages.Count);
        Assert.Equal(3, svgPages.Count);
        for (int index = 0; index < 3; index++) {
            Assert.True(OfficePngReader.TryDecode(pngPages[index].Bytes, out OfficeRasterImage? raster));
            Assert.True(raster!.GetPixel(20, index < 2 ? 20 : 5).R > 240);
            string svg = Encoding.UTF8.GetString(svgPages[index].Bytes);
            Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
            Assert.Contains("<image", svg, StringComparison.Ordinal);
        }
        Assert.Equal(3, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlImages_OversizedSvgContinuesAsVectorsThroughClippedPageFragments() {
        const string svgSource = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 90'><rect width='40' height='90' fill='red'/></svg>";
        string data = Convert.ToBase64String(Encoding.UTF8.GetBytes(svgSource));
        string html = "<html style='margin:0'><body style='margin:0'><img id='tall-vector' src='data:image/svg+xml;base64," + data
            + "' style='display:block;width:40px;height:90px'></body></html>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<OfficeImageExportResult> pngPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Png, options);
        IReadOnlyList<OfficeImageExportResult> svgPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Svg, options);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            HtmlRenderClipGroup fragment = Assert.Single(page.Visuals.OfType<HtmlRenderClipGroup>(), group => group.Source == "img#tall-vector");
            Assert.Single(fragment.Visuals.OfType<HtmlRenderDrawing>(), image => image.Source == "img#tall-vector");
        });
        for (int index = 0; index < 3; index++) {
            Assert.True(OfficePngReader.TryDecode(pngPages[index].Bytes, out OfficeRasterImage? raster));
            Assert.True(raster!.GetPixel(20, index < 2 ? 20 : 5).R > 240);
            string svg = Encoding.UTF8.GetString(svgPages[index].Bytes);
            Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
            Assert.Contains("<rect", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("data:image/svg+xml", svg, StringComparison.Ordinal);
        }
        Assert.Equal(3, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlImages_OversizedRoundedImageRetainsItsAuthoredPathClipAcrossFragments() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0));
        string html = "<html style='margin:0'><body style='margin:0'><img id='rounded-tall' src='data:image/png;base64," + data
            + "' style='display:block;width:40px;height:90px;border-radius:8px'></body></html>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<OfficeImageExportResult> pngPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Png, options);
        IReadOnlyList<OfficeImageExportResult> svgPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Svg, options);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            HtmlRenderClipGroup fragment = Assert.Single(page.Visuals.OfType<HtmlRenderClipGroup>(), group => group.Source == "img#rounded-tall:content-clip");
            Assert.Single(fragment.Visuals.OfType<HtmlRenderPathClipGroup>());
        });
        for (int index = 0; index < 3; index++) {
            Assert.True(OfficePngReader.TryDecode(pngPages[index].Bytes, out OfficeRasterImage? raster));
            Assert.True(raster!.GetPixel(20, index < 2 ? 20 : 5).R > 240);
            Assert.True(CountBackgroundOccurrences(Encoding.UTF8.GetString(svgPages[index].Bytes), "<clipPath") >= 2);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlImages_InvalidValuesAndRoundedClipUseSharedPathAndCatalogedDiagnostics() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(20, 10));
        string html = $"<img id='invalid-image' src='data:image/png;base64,{data}' style='display:block;width:30px;height:20px;object-fit:stretch;object-position:sideways;aspect-ratio:0/1'>"
            + $"<img id='rounded-image' src='data:image/png;base64,{data}' style='display:block;width:30px;height:20px;border-radius:8px 2px / 3px 6px'>";

        var options = new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 45D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlDiagnostic replaced = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ReplacedElementValueUnsupported);
        HtmlRenderPathClipGroup rounded = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderPathClipGroup>(), group => group.Source == "img#rounded-image:content-clip");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

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
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
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
        var options = new HtmlRenderOptions {
            ViewportWidth = 35D,
            ViewportHeight = 25D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderPathClipGroup group = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderPathClipGroup>(), item => item.Source == "div#rounded-background:background-image:clip");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(35D / HtmlRenderOptions.CssPixelsPerInch, 25D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(OfficeClipPathKind.Path, group.ClipPath.Kind);
        Assert.True(group.ClipPath.Commands.Count >= 10);
        Assert.Single(group.Visuals.OfType<HtmlRenderImagePattern>());
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(0, 0));
        Assert.True(raster.GetPixel(15, 10).A > 0);
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("<path", svg, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }
}
