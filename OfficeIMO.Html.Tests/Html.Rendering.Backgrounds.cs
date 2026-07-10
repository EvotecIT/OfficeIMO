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
            + "');background-size:cover;background-repeat:repeat\"></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 160D,
            Margins = HtmlRenderMargins.All(8D)
        });

        HtmlRenderImage background = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(100D, background.Width, 3);
        Assert.Equal(50D, background.Height, 3);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
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
}
