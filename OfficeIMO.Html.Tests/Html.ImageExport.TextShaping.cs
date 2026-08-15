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
    public void HtmlRenderer_ConfiguredShaperUsesTheActualFallbackRuns() {
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
        Assert.Contains(provider.Requests, request => request.Text == syriac && request.FontName == ManagedTextShapingTestAssets.FamilyName);
        Assert.Contains(provider.Requests, request => request.Text.Contains("Latin", StringComparison.Ordinal) && request.FontName == "Latin Test");
    }

    [Fact]
    public void HtmlRenderer_ConfiguredShaperAdvancesDriveInlineLayout() {
        const string text = "ܫܠܡ";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"margin:0;font-size:20px;line-height:20px;font-family:'OfficeIMO Shaping Test'\">" + text + "</p>");
        var provider = new FixedAdvanceTextShapingProvider(100);
        var options = new HtmlRenderOptions {
            ViewportWidth = 20D,
            Margins = HtmlRenderMargins.All(0D),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = provider,
            TextShapingLanguage = "syr"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(text.Select(character => (int)character).ToArray()));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderText[] visuals = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.Single(visuals);
        Assert.Equal(6D, visuals.Sum(visual => visual.TextAdvanceWidth ?? 0D), 3);
        Assert.Single(visuals.Select(visual => visual.Y).Distinct());
        Assert.Contains(provider.Requests, request => request.Text == text);
        Assert.Empty(rendered.Diagnostics);
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

    private sealed class FixedAdvanceTextShapingProvider : IOfficeTextShapingProvider {
        private readonly int _advanceWidth;

        internal FixedAdvanceTextShapingProvider(int advanceWidth) {
            _advanceWidth = advanceWidth;
        }

        internal List<OfficeTextShapingRequest> Requests { get; } = new List<OfficeTextShapingRequest>();

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Requests.Add(request);
            var glyphs = new List<OfficeShapedGlyph>();
            int textIndex = 0;
            foreach (string element in OfficeTextElements.Enumerate(request.Text)) {
                glyphs.Add(new OfficeShapedGlyph(1, element, textIndex, advanceWidth: _advanceWidth));
                textIndex += element.Length;
            }
            return new OfficeTextShapingResult(glyphs);
        }
    }
}
