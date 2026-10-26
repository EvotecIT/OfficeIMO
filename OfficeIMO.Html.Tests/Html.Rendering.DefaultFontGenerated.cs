using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using System.Net;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("<ol><li></li></ol>")]
    [InlineData("<ol style='font-family:Supplied,Arial'><li>A</li></ol>")]
    [InlineData("<ol><li><span style='font-family:Supplied'>A</span></li></ol>")]
    public void HtmlDefaultFont_GeneratedNumericMarkerRetainsItsPermittedDefaultFace(string body) {
        if (OfficeTrueTypeFont.TryLoadFontFamilyForText("Arial", OfficeFontFaceDescriptor.Regular, "1.", out _) == null) return;
        var options = new HtmlRenderOptions { Margins = HtmlRenderMargins.All(0D), ViewportWidth = 240D };
        options.Fonts.Add("Supplied", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A'));

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(body), options);
        HtmlRenderText marker = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>(), text => text.Text.StartsWith("1.", StringComparison.Ordinal));

        Assert.True(rendered.Fonts.TryResolveFaceForText(marker.Text, marker.Font.FamilyName,
            marker.FontDescriptor, marker.Font.Size, out OfficeFontFace? face));
        Assert.Equal(face!.Program.Measure(marker.Text, marker.Font.Size), marker.TextAdvanceWidth!.Value, 6);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated
            && item.Source == marker.Source);
    }

    [Fact]
    public void HtmlDefaultFont_CoveredListBodyAndMarkerDoNotDiscoverAnUnusedHostFace() {
        var options = new HtmlRenderOptions();
        options.Fonts.Add("Supplied", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', '1', '.', ' '));
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<ol style='font-family:Supplied,Arial'><li>A</li></ol>"), options);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == "InstalledFontResolved");
        Assert.DoesNotContain(rendered.Fonts.Faces, face => face.FamilyName == "Arial");
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlDefaultFont_MissingImageAltRetainsItsPermittedDefaultFace() {
        if (OfficeTrueTypeFont.TryLoadFontFamilyForText("Arial", OfficeFontFaceDescriptor.Regular, "ALT", out _) == null) return;
        var options = new HtmlRenderOptions {
            BaseUri = new Uri("https://images.example.test/"),
            ResourceResolver = (_, _) => System.Threading.Tasks.Task.FromResult<HtmlResolvedResource?>(null)
        };
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse("<img src='missing.png' alt='ALT'>"), options);
        HtmlRenderText alt = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>(), text => text.Text == "ALT");
        Assert.True(rendered.Fonts.TryResolveFaceForText(alt.Text, alt.Font.FamilyName,
            alt.FontDescriptor, alt.Font.Size, out OfficeFontFace? face));
        Assert.NotNull(face);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated
            && item.Source == alt.Source);
    }

    [Fact]
    public void HtmlDefaultFont_LaterFrameFontCannotEraseAnEarlierParentEstimate() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A');
        string child = "<style>@font-face{font-family:Late;src:url(data:font/ttf;base64,"
            + Convert.ToBase64String(font) + ")}</style><p style='font:16px Late'>A</p>";
        string html = "<p id='parent' style='font:16px Late'>AAAA</p><iframe srcdoc=\""
            + WebUtility.HtmlEncode(child) + "\"></iframe>";
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderText parent = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>(), text => text.Text == "AAAA");
        Assert.Contains(rendered.Fonts.Faces, face => face.FamilyName == "Late");
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated
            && diagnostic.Source == parent.Source);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }
}
