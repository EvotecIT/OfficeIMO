using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRenderer_ConfiguredShaperSatisfiesStrictJoiningScriptContract() {
        const string text = "ܫܠܡ";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'OfficeIMO Shaping Test'\">" + text + "</p>");
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = provider,
            TextShapingLanguage = "syr"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(text.Select(character => (int)character).ToArray()));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        OfficeImageExportResult image = document.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.DoesNotContain(image.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Contains(provider.Requests, request =>
            request.Text == text &&
            request.Language == "syr");
    }

    [Fact]
    public void HtmlRenderer_ConfiguredShaperProbesTheActualFallbackRun() {
        const string latin = "Latin ";
        const string syriac = "ܫܠܡ";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'Latin Test','OfficeIMO Shaping Test'\">" + latin + syriac + "</p>");
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new HtmlRenderOptions {
            ViewportWidth = 220D,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = provider,
            TextShapingLanguage = "syr"
        };
        options.Fonts.Add("Latin Test", ManagedTextShapingTestAssets.CreateFont(latin.Select(character => (int)character).ToArray()));
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(syriac.Select(character => (int)character).ToArray()));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Contains(provider.Requests, request => request.Text == syriac);
        Assert.DoesNotContain(provider.Requests, request => request.Text.Contains("Latin", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRasterExport_ReachesTheSharedTextShapingProvider() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'OfficeIMO Shaping Test'\">A</p>");
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeImageExportResult result = document.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
    }
}
