using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("ar", "fa")]
    [InlineData("fa", "ar")]
    [InlineData("ar", "AR")]
    [InlineData("ar", "ar")]
    public void HtmlRenderer_ConfiguredLanguageAdvancesKeepIndependentIntrinsicWidths(string first, string second) {
        HtmlRenderDocument rendered = RenderLanguageCacheFixture(new[] { first, second }, null, false,
            out LanguageAdvanceTextShapingProvider provider);
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).ToArray();
        string[] languages = { first, second };

        for (int index = 0; index < languages.Length; index++) {
            HtmlRenderShape box = Assert.Single(visuals.OfType<HtmlRenderShape>(), item => item.Source == "div#probe" + index);
            HtmlRenderText[] text = LanguageCacheText(visuals, index);
            double expected = 20D * LanguageAdvanceTextShapingProvider.Advance(languages[index]) / provider.Program.UnitsPerEm;
            Assert.Equal(expected, box.Width, 6);
            Assert.NotEmpty(text);
            Assert.Single(text.Select(item => item.Y).Distinct());
            Assert.Equal(expected, text.Sum(item => item.TextAdvanceWidth ?? 0D), 6);
            Assert.Contains(provider.Requests, request => request.Language == languages[index]);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("ar")]
    [InlineData("fa")]
    public void HtmlRenderer_GlobalLanguageOverrideControlsMeasuredAndPaintedAdvances(string language) {
        HtmlRenderDocument rendered = RenderLanguageCacheFixture(new[] { "ar", "fa" }, language, false,
            out LanguageAdvanceTextShapingProvider provider);
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).ToArray();
        double expected = 20D * LanguageAdvanceTextShapingProvider.Advance(language) / provider.Program.UnitsPerEm;

        for (int index = 0; index < 2; index++) {
            HtmlRenderShape box = Assert.Single(visuals.OfType<HtmlRenderShape>(), item => item.Source == "div#probe" + index);
            HtmlRenderText[] text = LanguageCacheText(visuals, index);
            Assert.Equal(expected, box.Width, 6);
            Assert.NotEmpty(text);
            Assert.Equal(expected, text.Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        }
        Assert.NotEmpty(provider.Requests);
        Assert.All(provider.Requests, request => Assert.Equal(language, request.Language));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlRenderer_ChangingLanguageAfterCachedTextUsesTheNewAdvanceForWrapping() {
        HtmlRenderDocument rendered = RenderLanguageCacheFixture(new[] { "ar", "fa" }, null, true,
            out LanguageAdvanceTextShapingProvider provider);
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).ToArray();
        HtmlRenderText[] text = LanguageCacheText(visuals, 1);
        double expectedWord = 20D * LanguageAdvanceTextShapingProvider.Advance("fa") / provider.Program.UnitsPerEm;

        Assert.NotEmpty(text);
        Assert.Equal(2, text.Select(item => item.Y).Distinct().Count());
        Assert.All(text.GroupBy(item => item.Y), line => Assert.Equal(expectedWord,
            line.Sum(item => item.TextAdvanceWidth ?? 0D), 6));
        Assert.Equal("بب", string.Concat(text.OrderBy(item => item.Y).Select(item => item.Text)).Replace(" ", string.Empty));
        Assert.Contains(provider.Requests, request => request.Text == "ب" && request.Language == "ar");
        Assert.Contains(provider.Requests, request => request.Text == "ب" && request.Language == "fa");
        rendered.RequireNoLoss();
    }

    private static HtmlRenderText[] LanguageCacheText(IEnumerable<HtmlRenderVisual> visuals, int index) =>
        visuals.OfType<HtmlRenderText>().Where(item => item.Color == (index == 0 ? OfficeColor.Red : OfficeColor.Blue)
            && item.Text.Trim().Length > 0).ToArray();

    private static HtmlRenderDocument RenderLanguageCacheFixture(string[] languages, string? globalLanguage, bool wrapSecond,
        out LanguageAdvanceTextShapingProvider provider) {
        // The layout path resolves both logical Arabic and presentation forms.
        // Use the same complete font as the public before/after reproduction;
        // a beh-only cmap would select a fallback context before our provider.
        byte[] font = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "Website", "Apps",
            "OfficeIMO.Web.Converter", "Assets", "Fonts", "NotoSansArabic-Regular.ttf"));
        provider = new LanguageAdvanceTextShapingProvider(font);
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, ViewportHeight = 300D,
            Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false,
            DefaultFontFamily = ManagedTextShapingTestAssets.FamilyName,
            TextShapingProvider = provider, TextShapingLanguage = globalLanguage
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, font);
        string html = "<style>html,body{margin:0;padding:0}.language-probe{display:inline-block;width:max-content;"
            + "margin:0;padding:0;border:0;font:20px/24px '" + ManagedTextShapingTestAssets.FamilyName
            + "';direction:ltr;background:#eeeeee}</style>";
        for (int index = 0; index < languages.Length; index++) {
            bool wrap = wrapSecond && index == 1;
            html += "<div class='language-probe' id='probe" + index + "' lang='" + languages[index]
                + "' style='color:" + (index == 0 ? "red" : "blue") + (wrap ? ";width:30px" : string.Empty)
                + "'>" + (wrap ? "ب ب" : "ب") + "</div><br>";
        }
        return HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
    }

    // A real decoded font plus this public provider makes language-sensitive
    // advances observable without depending on a particular font's locl table.
    private sealed class LanguageAdvanceTextShapingProvider : IOfficeTextShapingProvider {
        internal LanguageAdvanceTextShapingProvider(byte[] font) {
            Program = OfficeTrueTypeFont.TryLoad(font)!;
        }
        internal IOfficeFontProgram Program { get; }
        internal List<OfficeTextShapingRequest> Requests { get; } = new();
        internal static int Advance(string? language) => language == "ar" ? 400 : 1200;

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            request.CancellationToken.ThrowIfCancellationRequested();
            Requests.Add(request);
            if (request.Text.Length == 0) return null;
            var glyphs = new List<OfficeShapedGlyph>();
            for (int index = 0; index < request.Text.Length; index++) {
                char scalar = request.Text[index];
                if (!Program.TryGetGlyphMetrics(scalar, out int glyphId, out _)) return null;
                glyphs.Add(new OfficeShapedGlyph(glyphId, scalar.ToString(), index,
                    scalar == ' ' ? 100 : Advance(request.Language)));
            }
            if (request.Direction == OfficeTextDirection.RightToLeft) glyphs.Reverse();
            return new OfficeTextShapingResult(glyphs, request.Direction);
        }
    }
}
