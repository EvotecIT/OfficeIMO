using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public async Task HtmlRenderAsync_ResolvesExternalStylesheetBackgroundImageRelativeToTheStylesheet() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(12, 8);
        var requested = new List<string>();
        var options = new HtmlRenderOptions {
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(8D),
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                requested.Add(request.Uri.AbsoluteUri);
                if (request.Kind == HtmlResourceKind.Stylesheet) {
                    const string css = ".hero{width:120px;height:80px;background:#112233 url('../images/background.png') right bottom / 40px 20px no-repeat}";
                    return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css"));
                }

                Assert.Equal(HtmlResourceKind.Image, request.Kind);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(imageBytes, "image/png"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(
            "<link rel='stylesheet' href='https://assets.example.test/css/site.css'><div class='hero'>BackgroundMarker</div>",
            options);

        Assert.Equal(new[] {
            "https://assets.example.test/css/site.css",
            "https://assets.example.test/images/background.png"
        }, requested);
        HtmlRenderImage background = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(),
            image => image.Source != null && image.Source.EndsWith(":background-image", StringComparison.Ordinal));
        Assert.Equal(40D, background.Width, 3);
        Assert.Equal(20D, background.Height, 3);
        Assert.Equal(88D, background.X, 3);
        Assert.Equal(68D, background.Y, 3);
        Assert.Contains(
            "BackgroundMarker",
            string.Concat(rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(text => text.Text)),
            StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.StylesheetUrlResourcesPending);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public async Task HtmlRenderAsync_ResolvesEveryExternalBackgroundLayerRelativeToTheStylesheet() {
        byte[] red = PdfPngTestImages.CreateRgbPng(255, 0, 0);
        byte[] blue = PdfPngTestImages.CreateRgbPng(0, 0, 255);
        var requested = new List<string>();
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(8D),
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                requested.Add(request.Uri.AbsoluteUri);
                if (request.Kind == HtmlResourceKind.Stylesheet) {
                    const string css = ".hero{width:40px;height:40px;background:url('../images/top.png') left top / 10px 10px no-repeat,url('../images/bottom.png') left top / 20px 20px no-repeat}";
                    return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css"));
                }

                Assert.Equal(HtmlResourceKind.Image, request.Kind);
                return Task.FromResult<HtmlResolvedResource?>(request.Uri.AbsolutePath.EndsWith("/top.png", StringComparison.Ordinal)
                    ? new HtmlResolvedResource(red, "image/png")
                    : new HtmlResolvedResource(blue, "image/png"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(
            "<link rel='stylesheet' href='https://assets.example.test/css/site.css'><div class='hero'></div>",
            options);

        Assert.Equal(new[] {
            "https://assets.example.test/css/site.css",
            "https://assets.example.test/images/top.png",
            "https://assets.example.test/images/bottom.png"
        }, requested);
        IReadOnlyList<HtmlRenderImage> layers = rendered.Pages[0].Visuals.OfType<HtmlRenderImage>().ToList();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(2, layers.Count);
        Assert.EndsWith(":background-image[1]", layers[0].Source, StringComparison.Ordinal);
        Assert.EndsWith(":background-image[0]", layers[1].Source, StringComparison.Ordinal);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(8, 8));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(18, 18));
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.StylesheetUrlResourcesPending);
    }

    [Fact]
    public void HtmlRender_ValidatesAndClonesTheBackgroundLayerLimit() {
        var options = new HtmlRenderOptions { MaxBackgroundImageLayers = 7 };

        Assert.Equal(7, options.Clone().MaxBackgroundImageLayers);
        options.MaxBackgroundImageLayers = 0;
        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlRenderEngine.Render("<div></div>", options));
    }

    [Fact]
    public void HtmlBackgroundImage_FlowsThroughSharedPngSvgAndSearchablePdfBackends() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(6, 4));
        string html = "<div style=\"width:100px;height:60px;background-image:url('data:image/png;base64,"
            + imageData
            + "');background-repeat:no-repeat;background-size:30px 20px;background-position:right bottom\">BackgroundOutputMarker</div>";
        var imageOptions = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 180D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, imageOptions);
        HtmlRenderImage background = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        OfficeImageExportResult png = html.ExportImage(OfficeImageExportFormat.Png, imageOptions);
        OfficeImageExportResult svg = html.ExportImage(OfficeImageExportFormat.Svg, imageOptions);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        byte[] pdf = html.SaveAsPdf(pdfOptions);

        Assert.EndsWith(":background-image", background.Source, StringComparison.Ordinal);
        Assert.Equal(30D, background.Width, 3);
        Assert.Equal(20D, background.Height, 3);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("data:image/png;base64", Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
        string pdfText = PdfCore.PdfReadDocument.Load(pdf).ExtractText().Replace("\r", string.Empty).Replace("\n", string.Empty);
        Assert.Contains("BackgroundOutputMarker", pdfText, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRender_DiagnosesDeterministicBackgroundFallbacks() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 1));
        string source = "data:image/png;base64," + imageData;
        string html = "<div style=\"width:100px;height:100px;background-image:url('"
            + source
            + "'),url('"
            + source
            + "');background-size:unsupported-size;background-repeat:space\"></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 160D,
            Margins = HtmlRenderMargins.All(8D)
        });

        IReadOnlyList<HtmlRenderImage> backgrounds = rendered.Pages[0].Visuals.OfType<HtmlRenderImage>().ToList();
        Assert.Equal(2, backgrounds.Count);
        Assert.All(backgrounds, background => Assert.Equal(100D, background.Width, 3));
        Assert.All(backgrounds, background => Assert.Equal(50D, background.Height, 3));
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlBackgroundLayers_PaintBackToFrontAcrossPngSvgAndPdf() {
        string red = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0));
        string blue = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(0, 0, 255));
        string html = "<div style=\"width:40px;height:40px;background-image:url('data:image/png;base64,"
            + red
            + "'),url('data:image/png;base64,"
            + blue
            + "');background-size:10px 10px,20px 20px;background-position:left top,left top;background-repeat:no-repeat,no-repeat\"></div>";
        var imageOptions = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, imageOptions);
        IReadOnlyList<HtmlRenderImage> layers = rendered.Pages[0].Visuals.OfType<HtmlRenderImage>().ToList();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = html.ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        byte[] pdf = html.SaveAsPdf(pdfOptions);

        Assert.Equal(2, layers.Count);
        Assert.EndsWith(":background-image[1]", layers[0].Source, StringComparison.Ordinal);
        Assert.EndsWith(":background-image[0]", layers[1].Source, StringComparison.Ordinal);
        Assert.Equal(20D, layers[0].Width, 3);
        Assert.Equal(10D, layers[1].Width, 3);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(8, 8));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(18, 18));
        Assert.Equal(2, CountBackgroundOccurrences(svg, "data:image/png;base64,"));
        Assert.Equal(2, PdfCore.PdfImageExtractor.ExtractImagePlacements(pdf).Count);
        Assert.Equal(2, PdfCore.PdfImageExtractor.ExtractImages(pdf).Count);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit);
    }

    [Fact]
    public void HtmlRender_BoundsBackgroundLayersAndDiagnosesOmittedGradients() {
        string red = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0));
        string blue = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(0, 0, 255));
        string green = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(0, 255, 0));
        string html = "<div style=\"width:40px;height:40px;background-image:linear-gradient(red,blue),url('data:image/png;base64,"
            + red
            + "'),url('data:image/png;base64,"
            + blue
            + "'),url('data:image/png;base64,"
            + green
            + "');background-repeat:no-repeat\"></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D),
            MaxBackgroundImageLayers = 2
        });

        Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlRender_ClipsCoverBackgroundToThePaintArea() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 1));
        string html = "<div style=\"width:100px;height:100px;background-image:url('data:image/png;base64,"
            + imageData
            + "');background-position:center;background-size:cover;background-repeat:no-repeat\"></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 140D,
            Margins = HtmlRenderMargins.All(8D)
        });

        HtmlRenderImage background = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(100D, background.Width, 3);
        Assert.Equal(100D, background.Height, 3);
        Assert.Equal(0.25D, background.SourceCrop.Left, 3);
        Assert.Equal(0.25D, background.SourceCrop.Right, 3);
        Assert.Equal(0D, background.SourceCrop.Top, 3);
        Assert.Equal(0D, background.SourceCrop.Bottom, 3);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlBackgroundRepeat_UsesOneBoundedPatternAcrossPngSvgAndPdf() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 2));
        string html = "<div style=\"width:20px;height:10px;background-image:url('data:image/png;base64,"
            + imageData
            + "');background-size:8px 4px;background-repeat:repeat\"></div>";
        var imageOptions = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 60D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, imageOptions);
        HtmlRenderImagePattern pattern = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImagePattern>());
        OfficeDrawing drawing = rendered.Pages[0].CreateDrawing();
        OfficeImageExportResult png = html.ExportImage(OfficeImageExportFormat.Png, imageOptions);
        string svg = html.ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        byte[] pdf = html.SaveAsPdf(pdfOptions);

        Assert.Equal(9L, pattern.Pattern.EstimatedTileCount);
        Assert.Single(drawing.ImagePatterns);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("<pattern", svg, StringComparison.Ordinal);
        Assert.Equal(1, CountBackgroundOccurrences(svg, "data:image/png;base64,"));
        Assert.Equal(9, PdfCore.PdfImageExtractor.ExtractImagePlacements(pdf).Count);
        Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRender_BoundsOperationWideBackgroundTileExpansion() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        string html = "<div style=\"width:100px;height:100px;background-image:url('data:image/png;base64,"
            + imageData
            + "');background-size:1px 1px;background-repeat:repeat\"></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 140D,
            Margins = HtmlRenderMargins.All(8D),
            MaxBackgroundImageTiles = 8
        });

        Assert.Empty(rendered.Pages[0].Visuals.OfType<HtmlRenderImagePattern>());
        Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded);
    }

    [Fact]
    public void HtmlRender_PropagatesRootBackgroundToTheSurfaceBehindContent() {
        const string html = "<style>body{background-color:#123456}</style><p>RootCanvasMarker</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 160D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderShape surface = Assert.IsType<HtmlRenderShape>(page.Visuals[0]);
        HtmlRenderShape rootBackground = Assert.IsType<HtmlRenderShape>(page.Visuals[1]);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        Assert.Equal("render-surface", surface.Source);
        Assert.Equal("render-root-background", rootBackground.Source);
        Assert.Equal(int.MinValue, surface.PaintOrder);
        Assert.Equal(int.MinValue + 1, rootBackground.PaintOrder);
        Assert.Contains(page.Visuals.Skip(2), visual => visual is HtmlRenderText text && text.Text.Contains("RootCanvas", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.FromRgb(0x12, 0x34, 0x56), raster.GetPixel(raster.Width - 1, raster.Height - 1));
    }

    [Fact]
    public void HtmlRender_RootBackgroundDoesNotCreateAFalseBlankPageBeforeFirstBreak() {
        const string html = "<style>body{background:#f0f0f0}</style><p style='break-before:page'>FirstPageMarker</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(3D, 2D),
            Margins = HtmlRenderMargins.All(12D)
        });

        HtmlRenderPage page = Assert.Single(rendered.Pages);
        Assert.Contains(page.Visuals, visual => visual.Source == "render-root-background");
        Assert.Contains(
            "FirstPageMarker",
            string.Concat(page.Visuals.OfType<HtmlRenderText>().Select(text => text.Text)),
            StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_PrefersDocumentRootBackgroundForCanvasPropagation() {
        const string html = "<style>html{background:#654321}body{background:#123456}</style><p>DocumentRootMarker</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 160D,
            Margins = HtmlRenderMargins.All(8D)
        });

        HtmlRenderShape rootBackground = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "render-root-background");
        Assert.Equal(OfficeColor.FromRgb(0x65, 0x43, 0x21), rootBackground.Shape.FillColor);
    }

    [Fact]
    public void HtmlRender_PaintsBackgroundImagesOnTableCells() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 4));
        string html = "<table style='width:120px'><tr><td style=\"height:40px;background-image:url('data:image/png;base64,"
            + imageData
            + "');background-repeat:no-repeat;background-size:16px 16px;background-position:center\">CellMarker</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 180D,
            Margins = HtmlRenderMargins.All(8D)
        });

        HtmlRenderImage background = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.EndsWith(":background-image", background.Source, StringComparison.Ordinal);
        Assert.Equal(16D, background.Width, 3);
        Assert.Equal(16D, background.Height, 3);
        Assert.Contains(
            "CellMarker",
            string.Concat(rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(text => text.Text)),
            StringComparison.Ordinal);
    }

    private static int CountBackgroundOccurrences(string value, string marker) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(marker, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += marker.Length;
        }

        return count;
    }
}
