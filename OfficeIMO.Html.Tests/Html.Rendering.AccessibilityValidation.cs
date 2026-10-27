using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdfAccessibilityValidator_UsesScreenMediaForSourceImageNames() {
        string raster = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<html lang='en'><style>@media print {img{display:none}}</style><body><p>Screen report</p>" +
            "<img width='12' height='12' src='data:image/png;base64," + raster + "'></body></html>");
        HtmlPdfRenderRequestResult rendered = source.RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.Pdf));

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(rendered);

        Assert.Contains(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityImageNameMissing");
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_UsesCallerStylesheetsForSourceImageNames() {
        string raster = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<html lang='en'><body><p>Print report</p><img width='12' height='12' src='data:image/png;base64," + raster + "'></body></html>");
        var options = new HtmlToPdfOptions();
        options.AdditionalStylesheets.Add("img {display:none}");
        HtmlPdfRenderRequestResult rendered = source.RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(rendered);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues.Select(issue => issue.Message)));
    }

    [Theory]
    [InlineData(480, false)]
    [InlineData(960, true)]
    public async Task HtmlPdfAccessibilityValidator_RetainsResolvedResponsiveStylesWithoutFetchingAgain(int viewportWidth, bool valid) {
        int resourceCalls = 0;
        string raster = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<html lang='en'><head><link rel='stylesheet' href='https://assets.example.test/report.css'></head>" +
            "<body><p>Responsive report</p><img width='12' height='12' src='data:image/png;base64," + raster + "'></body></html>");
        var options = new HtmlToPdfOptions {
            ViewportWidth = viewportWidth,
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreateTrustedHost(),
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                resourceCalls++;
                Assert.Equal(HtmlResourceKind.Stylesheet, request.Kind);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    Encoding.UTF8.GetBytes("@media screen and (min-width:700px) {img {display:none}}"), "text/css"));
            }
        };
        HtmlPdfRenderRequestResult rendered = await source.RenderToPdfResultAsync(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, options));
        Assert.Equal(1, resourceCalls);
        // A completed operation must not depend on the caller's later configuration edits.
        options.AdditionalStylesheets.Add("img {display:none}");
        options.ViewportWidth = 1200;
        options.ResourceResolver = (_, _) => throw new InvalidOperationException("Validation must not resolve resources again.");

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(rendered);

        Assert.Equal(valid, result.IsValid);
        Assert.Equal(!valid, result.Issues.Any(issue => issue.Code == "HtmlPdfAccessibilityImageNameMissing"));
        Assert.Equal(1, resourceCalls);
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_AcceptsCompleteGeneratedStructure() {
        string rasterData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = """
            <!doctype html>
            <html lang="en-US">
              <body>
                <main>
                  <h1>Accessibility contract</h1>
                  <p>Paragraph with an <a href="https://example.test/accessibility">accessible link</a>.</p>
                  <ol><li>First item</li><li>Second item</li></ol>
                  <table>
                    <caption>Status</caption>
                    <tr><th scope="col">Area</th><th scope="col">State</th></tr>
                    <tr><td>Renderer</td><td>Ready</td></tr>
                  </table>
                  <img alt="Green status badge" width="24" height="24" src="data:image/png;base64,RASTER_DATA">
                </main>
              </body>
            </html>
            """.Replace("RASTER_DATA", rasterData);

        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        byte[] pdf = (source.ToPdf(new HtmlToPdfOptions())).ToBytes();

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(source, pdf);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues.Select(issue => $"{issue.Code}: {issue.Message}")));
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_ReportsMissingLanguageAndFigureAlt() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<main><h1>Missing metadata</h1><img width='12' height='12' src='data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1)) + "'></main>");
        byte[] pdf = (source.ToPdf(new HtmlToPdfOptions())).ToBytes();

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(source, pdf);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityLanguageMissing");
        Assert.Contains(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityImageNameMissing");
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_ReportsExplicitlyUntaggedOutput() {
        var options = new HtmlToPdfOptions {
            PdfOptions = new OfficeIMO.Pdf.PdfCreateOptions {
                Language = "en-US",
                TaggedStructureMode = OfficeIMO.Pdf.PdfTaggedStructureMode.None
            }
        };
        byte[] pdf = (HtmlConversionDocument.Parse("<p>Untagged output</p>").ToPdf(options)).ToBytes();

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(pdf);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityTagsMissing");
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_IgnoresImagesExcludedFromRenderedAccessibilityOutput() {
        string raster = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        string html = """
            <!doctype html><html lang="en"><head><style>.removed{display:none}.invisible{visibility:hidden}</style></head><body>
              <img hidden src="data:image/png;base64,RASTER">
              <div aria-hidden="true"><img src="data:image/png;base64,RASTER"></div>
              <div class="removed"><img src="data:image/png;base64,RASTER"></div>
              <img class="invisible" src="data:image/png;base64,RASTER">
              <img role="presentation" src="data:image/png;base64,RASTER">
              <p>Visible content</p>
            </body></html>
            """.Replace("RASTER", raster);
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(source, (source.ToPdf(new HtmlToPdfOptions())).ToBytes());

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues.Select(issue => $"{issue.Code}: {issue.Message}")));
        Assert.DoesNotContain(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityImageNameMissing");
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_StillRequiresNamesForAriaHiddenImageInputs() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<!doctype html><html lang='en'><body><div aria-hidden='true'><input type='image'></div><p>Visible content</p></body></html>");

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(source, (source.ToPdf(new HtmlToPdfOptions())).ToBytes());

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityImageNameMissing");
    }
}
