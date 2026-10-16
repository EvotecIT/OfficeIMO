using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderEncoder.Png, null)]
    [InlineData(HtmlRenderEncoder.Svg, null)]
    [InlineData(HtmlRenderEncoder.Png, " ")]
    [InlineData(HtmlRenderEncoder.Svg, " ")]
    public void HtmlArchive_RetainsPdfOmissionWhenSourceDiagnosticsAreInformational(HtmlRenderEncoder encoder, string? source) {
        const string html = "<style>.icon::before{font-family:MissingIcon;content:'\\F42B'}</style>"
            + "<p class='icon'>Voyager Overview</p>";
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        };
        HtmlRenderResult retained = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, encoder, options))
            .WithAdditionalDiagnostics(new[] {
                new HtmlDiagnostic("source", "source-info", "Source was inspected.", HtmlDiagnosticSeverity.Info, source)
            });
        OfficeImageExportResult image = Assert.Single(retained.ExportImages());
        HtmlRenderArchiveResult archive = retained.ExportArchive();
        HtmlRenderArchivePage page = Assert.Single(archive.Manifest.Pages);

        Assert.Contains(image.Diagnostics, diagnostic => diagnostic.Code == HtmlPdfDiagnosticCodes.UnavailablePrivateUseGlyphOmitted
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(page.EncodingDiagnostics, diagnostic => diagnostic.Code == HtmlPdfDiagnosticCodes.UnavailablePrivateUseGlyphOmitted
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain(page.EncodingDiagnostics, diagnostic => diagnostic.Code == "source-info");
        Assert.True(page.HasLoss);
        Assert.True(archive.Manifest.HasLoss);
        Assert.Contains(archive.Manifest.Diagnostics, diagnostic => diagnostic.Code == "source-info");
    }
}
