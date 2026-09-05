using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
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
        byte[] pdf = source.ToPdfBytes(new HtmlToPdfOptions());

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(source, pdf);

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues.Select(issue => $"{issue.Code}: {issue.Message}")));
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_ReportsMissingLanguageAndFigureAlt() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<main><h1>Missing metadata</h1><img width='12' height='12' src='data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1)) + "'></main>");
        byte[] pdf = source.ToPdfBytes(new HtmlToPdfOptions());

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(source, pdf);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityLanguageMissing");
        Assert.Contains(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityImageNameMissing");
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_ReportsExplicitlyUntaggedOutput() {
        var options = new HtmlToPdfOptions {
            PdfOptions = new OfficeIMO.Pdf.PdfOptions {
                Language = "en-US",
                TaggedStructureMode = OfficeIMO.Pdf.PdfTaggedStructureMode.None
            }
        };
        byte[] pdf = HtmlConversionDocument.Parse("<p>Untagged output</p>").ToPdfBytes(options);

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

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(source, source.ToPdfBytes(new HtmlToPdfOptions()));

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues.Select(issue => $"{issue.Code}: {issue.Message}")));
        Assert.DoesNotContain(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityImageNameMissing");
    }

    [Fact]
    public void HtmlPdfAccessibilityValidator_StillRequiresNamesForAriaHiddenImageInputs() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<!doctype html><html lang='en'><body><div aria-hidden='true'><input type='image'></div><p>Visible content</p></body></html>");

        HtmlPdfAccessibilityValidationResult result = HtmlPdfAccessibilityValidator.Validate(source, source.ToPdfBytes(new HtmlToPdfOptions()));

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Code == "HtmlPdfAccessibilityImageNameMissing");
    }
}
