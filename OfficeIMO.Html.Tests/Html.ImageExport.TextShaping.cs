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

    [Theory]
    [InlineData("<p>\u0915</p>", "\u0915", true)]
    [InlineData("<style>p::first-letter{font-size:16px}</style><p style='font-size:0'>\u0915hidden</p>", "\u0915", true)]
    [InlineData("<style>p::before{content:'\u0915hidden';font-size:0}p::first-letter{font-size:16px}</style><p>A</p>", "\u0915A", true)]
    [InlineData("<style>p::first-letter{font-size:16px}</style><p style='font-size:0'>A\u0915</p>", "A", false)]
    [InlineData("<style>p::before{content:'A\u0915';font-size:0}p::first-letter{font-size:16px}</style><p>A</p>", "AA", false)]
    public void HtmlRenderer_ReportsWhenConfiguredProviderDeclinesShapingRequiredScript(string content, string paintedText, bool expectedLoss) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>p{font-family:'OfficeIMO Shaping Test'}</style>" + content);
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            TextShapingProvider = OfficeManagedTextShapingProvider.Instance,
            TextShapingLanguage = "hi"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A', '\u0915'));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.Equal(paintedText, string.Concat(FirstLetterText(rendered).Select(text => text.Text)));
        if (!expectedLoss) {
            Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
            rendered.RequireNoLoss();
            options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
            HtmlRenderTestDriver.Render(document, options).RequireNoLoss();
            return;
        }

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item =>
            item.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Equal("provider-declined", diagnostic.Detail);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());

        options.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        HtmlConversionException strict = Assert.Throws<HtmlConversionException>(() =>
            HtmlRenderTestDriver.Render(document, options));
        Assert.Contains(strict.Diagnostics, item =>
            item.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void HtmlRenderer_ReportsUnpositionedMarksBeforePaintedGlyphExport(bool explicitManaged, bool explicitFeatures) {
        string features = explicitFeatures ? "font-feature-settings:'kern' 1;" : string.Empty;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"" + features + "font-family:'OfficeIMO Shaping Test'\">e\u0301 \u0627\u064E</p>");
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            TextShapingProvider = explicitManaged ? OfficeManagedTextShapingProvider.Instance : null
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(' ', 'e', 0x0301, 0x0627, 0xFE8D, 0x064E));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Approximation
            && (diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
                || diagnostic.Code == HtmlRenderDiagnosticCodes.OpenTypeFeatureUnsupported));
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlRenderer_ConfiguredProviderCanAcceptCombiningMarks() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'OfficeIMO Shaping Test'\">e\u0301</p>");
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = new ManagedTextShapingTestAssets.RecordingProvider()
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, ManagedTextShapingTestAssets.CreateFont('e', 0x0301));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        rendered.RequireNoLoss();
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

    [Fact]
    public void HtmlRenderer_ShapedMeasurementCacheRetainsOnlyBoundedSuccessfulShortKeys() {
        var cache = new HtmlShapedTextMeasurementCache();
        var font = new OfficeFontInfo("OfficeIMO Shaping Test", 12D);

        for (int index = 0; index < HtmlShapedTextMeasurementCache.MaximumEntries * 2; index++) {
            cache.Store("token-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture), font,
                OfficeTextFeatureSettings.Default, "ar", index);
        }
        string oversizedText = new string('x', HtmlShapedTextMeasurementCache.MaximumCacheableTextCharacters + 1);
        cache.Store(oversizedText, font, OfficeTextFeatureSettings.Default, "ar", 1D);

        Assert.Equal(HtmlShapedTextMeasurementCache.MaximumEntries, cache.Count);
        Assert.InRange(cache.RetainedKeyCharacters, 1, HtmlShapedTextMeasurementCache.MaximumRetainedKeyCharacters);
        Assert.True(cache.TryGet("token-0", font, OfficeTextFeatureSettings.Default, "ar", out double measured));
        Assert.Equal(0D, measured);
        Assert.False(cache.TryGet(oversizedText, font, OfficeTextFeatureSettings.Default, "ar", out _));

        // Language is authored input too; a short text cannot retain an unbounded
        // language key, and successful repeated stores must not consume its budget.
        var languageCache = new HtmlShapedTextMeasurementCache();
        string language = new string('x', HtmlShapedTextMeasurementCache.MaximumRetainedKeyCharacters - 1);
        languageCache.Store("A", font, OfficeTextFeatureSettings.Default, language, 8D);
        languageCache.Store("A", font, OfficeTextFeatureSettings.Default, language, 8D);
        languageCache.Store("B", font, OfficeTextFeatureSettings.Default, "ar", 8D);
        Assert.Equal(1, languageCache.Count);
        Assert.Equal(HtmlShapedTextMeasurementCache.MaximumRetainedKeyCharacters, languageCache.RetainedKeyCharacters);
        Assert.True(languageCache.TryGet("A", font, OfficeTextFeatureSettings.Default, language, out _));
        var rejected = new HtmlShapedTextMeasurementCache();
        rejected.Store("A", font, OfficeTextFeatureSettings.Default, language + "x", 8D);
        Assert.Equal(0, rejected.Count);
        Assert.Equal(0, rejected.RetainedKeyCharacters);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlRenderer_ReportsUnfulfilledRequiredArabicLigatures(bool explicitManaged) {
        const string text = "\u0644\u0627";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'OfficeIMO Shaping Test'\">" + text + "</p>");
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            TextShapingProvider = explicitManaged ? OfficeManagedTextShapingProvider.Instance : null
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(0x0644, 0x0627, 0xFEDF, 0xFE8E));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item =>
            item.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Equal("provider-declined", diagnostic.Detail);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlRenderer_ConfiguredProviderCanOwnRequiredArabicLigatures() {
        const string text = "\u0644\u0627";
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'OfficeIMO Shaping Test'\">" + text + "</p>");
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = provider
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(0x0644, 0x0627, 0xFEDF, 0xFE8E));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.Contains(provider.Requests, request => request.Text == text);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        rendered.RequireNoLoss();
        // This proves provider acceptance, not actual Arabic glyph/cluster fidelity.
        // Use the already-shipped optional HarfBuzz owner and exact native font for visual qualification.
    }

    [Fact]
    public void HtmlRenderer_HonorsExplicitDisabledRequiredLigatures() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-feature-settings:'rlig' 0;font-family:'OfficeIMO Shaping Test'\">\u0644\u0627</p>");
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextShapingProvider = OfficeManagedTextShapingProvider.Instance
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(0x0644, 0x0627, 0xFEDF, 0xFE8E));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.OpenTypeFeatureUnsupported);
        rendered.RequireNoLoss();
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
