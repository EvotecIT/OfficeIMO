using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("سلام بالعالم", "ar", "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf", false)]
    [InlineData("السَّلَامُ عَلَيْكُمْ", "ar", "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf", false)]
    [InlineData("שָׁלוֹם עוֹלָם", "he", "OfficeIMO.Drawing.Tests/TestAssets/OpenSans-Bold.ttf", true)]
    [InlineData("Cafe\u0301 A\u0301 a\u0323\u0301", "fr", "OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf", false)]
    public void HtmlManagedTypography_DefaultRouteRetainsPositionedGlyphsAndLogicalText(
        string text, string language, string fontPath, bool bold) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:360px 240px;margin:20px}html,body,p{margin:0}</style>"
            + "<p lang='" + language + "' dir='" + (language is "ar" or "he" ? "rtl" : "ltr")
            + "' style='font-family:Typography;font-size:24px;line-height:1.5;font-weight:"
            + (bold ? "700" : "400") + "'>" + text + "</p>");
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        options.Fonts.Add("Typography", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), fontPath)),
            bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.Equal(text, rendered.Text.Trim());
        rendered.RequireNoLoss();
        HtmlRenderText[] runs = Assert.Single(rendered.Pages).Visuals.OfType<HtmlRenderText>()
            .Where(run => !string.IsNullOrWhiteSpace(run.Text)).ToArray();
        Assert.NotEmpty(runs);
        Assert.Contains(runs, run => run.PreparedShaping != null);
        Assert.All(runs.Where(run => run.PreparedShaping != null), run => {
            OfficePreparedTextShaping prepared = run.PreparedShaping!;
            Assert.True(prepared.Measure(run.Font.Size) > 0D);
            Assert.Equal(prepared.Text, string.Concat(prepared.Result.Glyphs
                .Where(glyph => glyph.UnicodeText.Length > 0).OrderBy(glyph => glyph.TextIndex)
                .Select(glyph => glyph.UnicodeText)));
        });
    }

    [Fact]
    public void HtmlManagedTypography_DefaultRouteKeepsJoiningAcrossPaintOnlyBoundaries() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:240px 160px;margin:0}html,body,p{margin:0}</style>"
            + "<p lang='ar' dir='rtl' style='font:24px Typography'><span style='color:red'>م</span>"
            + "<span style='color:blue'>رحبا</span></p>");
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        options.Fonts.Add("Typography", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf")));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.Equal("مرحبا", rendered.Text.Trim());
        rendered.RequireNoLoss();
        HtmlRenderText[] runs = Assert.Single(rendered.Pages).Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Contains(runs, run => run.Color == OfficeColor.Red && run.PreparedShaping != null);
        Assert.Contains(runs, run => run.Color == OfficeColor.Blue && run.PreparedShaping != null);
    }
    [Fact]
    public void HtmlManagedTypography_MixedParagraphRetainsBothDirectionalMarkRuns() {
        const string source = "Cafe\u0301 שָׁלוֹם A\u0301";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:360px 240px;margin:20px}html,body,p{margin:0}</style>"
            + "<p dir='ltr' style='font:700 22px Typography'>" + source + "</p>");
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        options.Fonts.Add("Typography", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.Drawing.Tests/TestAssets/OpenSans-Bold.ttf")), OfficeFontStyle.Bold);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.Equal(source, rendered.Text.Trim());
        rendered.RequireNoLoss();
        HtmlRenderText[] marked = Assert.Single(rendered.Pages).Visuals.OfType<HtmlRenderText>()
            .Where(run => OfficeTextElements.ContainsCombiningMark(run.Text)).ToArray();
        Assert.NotEmpty(marked);
        Assert.All(marked, run => Assert.NotNull(run.PreparedShaping));
    }

    [Theory]
    [InlineData("Cafe\u0301", "ltr", "OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf", "letter-spacing:.5px")]
    [InlineData("Cafe\u0301 Cafe\u0301", "ltr", "OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf", "word-spacing:2px")]
    [InlineData("السَّلَامُ", "rtl", "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf", "letter-spacing:.5px")]
    public void HtmlManagedTypography_UnretainedSpacingCannotClaimNoLoss(
        string source, string direction, string fontPath, string spacing) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:360px 240px;margin:20px}html,body,p{margin:0}</style>"
            + "<p dir='" + direction + "' style='font:24px Typography;" + spacing + "'>" + source + "</p>");
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        options.Fonts.Add("Typography", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), fontPath)));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.Equal(source, rendered.Text.Trim());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

}
