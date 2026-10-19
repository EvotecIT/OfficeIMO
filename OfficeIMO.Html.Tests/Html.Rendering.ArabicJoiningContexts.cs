using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string ArabicContextWord = "مرحبا";
    private const string ArabicContextForms = "\uFEE3\uFEAE\uFEA3\uFE92\uFE8E";

    [Fact]
    public void HtmlArabicJoining_ColorOnlyRunsRetainContextSourceAndLinks() {
        const string content = "<a href='https://example.com/arabic' style='text-decoration:none'>"
            + "<span id='left' style='color:red'>م</span><span id='right' style='color:blue'>رحبا</span></a>";
        HtmlRenderDocument rendered = RenderArabicContext(content);
        HtmlRenderText[] text = ArabicContextText(rendered).ToArray();
        HtmlRenderText meem = Assert.Single(text, item => item.Color == OfficeColor.Red);
        HtmlRenderText[] remainder = text.Where(item => item.Color == OfficeColor.Blue).ToArray();

        Assert.Equal("\uFEE3", meem.Text);
        Assert.Equal("\uFE8E\uFE92\uFEA3\uFEAE", string.Concat(remainder.OrderBy(item => item.X).Select(item => item.Text)));
        Assert.Equal(ArabicContextWord, rendered.Text.Trim());
        Assert.All(text, item => Assert.Equal("https://example.com/arabic", item.LinkUri));
        Assert.All(text, item => Assert.Equal(ManagedTextShapingTestAssets.FamilyName, item.Font.FamilyName));
        Assert.All(text, item => Assert.Equal("span", item.Source));
        Assert.True(meem.X > remainder.Max(item => item.X));
        Assert.NotNull(meem.LogicalTextOrder);
        Assert.All(remainder, item => Assert.True(item.LogicalTextOrder.GetValueOrDefault() > meem.LogicalTextOrder.GetValueOrDefault()));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-content", false)]
    [InlineData("max-content", false)]
    [InlineData("max-content", true)]
    public void HtmlArabicJoining_IntrinsicAndFirstLineWidthsUseThePaintedForms(string width, bool firstLine) {
        const string split = "<span style='color:red'>م</span><span style='color:blue'>رحبا</span>";
        string css = firstLine ? "#probe::first-line{color:green}" : string.Empty;
        HtmlRenderDocument rendered = RenderArabicContext(split, width, css);
        HtmlRenderDocument plain = RenderArabicContext(ArabicContextWord, width, css);
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ArabicContextFont())!;
        double expected = font.Measure(ArabicContextForms, 20D);

        Assert.NotEqual(font.Measure(ArabicContextWord, 20D), expected);
        Assert.Equal(expected, ArabicContextBox(rendered).Width, 6);
        Assert.Equal(expected, ArabicContextBox(plain).Width, 6);
        Assert.Equal(expected, ArabicContextText(rendered).Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        Assert.Equal(ArabicContextWord, rendered.Text.Trim());
        Assert.Equal(new string(ArabicContextForms.Reverse().ToArray()),
            string.Concat(ArabicContextText(rendered).OrderBy(item => item.X).Select(item => item.Text)));
        if (firstLine) Assert.All(ArabicContextText(rendered), item => Assert.Equal(OfficeColor.Green, item.Color));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(" ")]
    [InlineData("\u200C")]
    [InlineData("<br>")]
    [InlineData("<span style='display:inline-block;width:8px;height:8px'></span>")]
    public void HtmlArabicJoining_AuthoredAndAtomicBoundariesDoNotJoin(string separator) {
        HtmlRenderDocument rendered = RenderArabicContext("<span style='color:red'>م</span>" + separator
            + "<span style='color:blue'>رحبا</span>");
        HtmlRenderText[] text = ArabicContextText(rendered).ToArray();
        Assert.Contains(text, item => item.Text.Contains('\uFEE1'));
        Assert.Contains(text, item => item.Text.Contains('\uFEAD'));
        Assert.DoesNotContain(text, item => item.Text.Contains('\uFEE3') || item.Text.Contains('\uFEAE'));
        Assert.Equal(ArabicContextWord, OfficeArabicTextShaper.ToLogicalText(rendered.Text)
            .Replace(" ", string.Empty).Replace("\u200C", string.Empty).Replace("\r", string.Empty).Replace("\n", string.Empty).Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlArabicJoining_TransparentJoinerAndMarkRemainInLogicalSource() {
        const string logical = "م\u064E\u200Dرحبا";
        HtmlRenderDocument rendered = RenderArabicContext("<span style='color:red'>م\u064E\u200D</span>"
            + "<span style='color:blue'>رحبا</span>");
        string painted = string.Concat(ArabicContextText(rendered).Select(item => item.Text));
        Assert.Contains('\uFEE3', painted);
        Assert.Contains('\uFEAE', painted);
        Assert.Equal(logical, rendered.Text.Trim());
        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    private static HtmlRenderDocument RenderArabicContext(string content, string width = "300px", string extraCss = "") {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false,
            DefaultFontFamily = ManagedTextShapingTestAssets.FamilyName
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, ArabicContextFont());
        string html = "<style>html,body{margin:0;padding:0}#probe{margin:0;padding:0;border:0;"
            + "font:20px/24px '" + ManagedTextShapingTestAssets.FamilyName + "';"
            + "background:white;display:inline-block;direction:rtl;width:" + width + "}"
            + extraCss + "</style><div id='probe'>" + content + "</div>";
        return HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
    }

    private static byte[] ArabicContextFont() {
        var advances = new Dictionary<int, ushort>();
        foreach (char scalar in ArabicContextWord + " \u064E\u200C\u200D") advances[scalar] = 700;
        foreach (char scalar in "\uFEE1\uFEAD\uFEA3\uFE92\uFE8E") advances[scalar] = 500;
        foreach (char scalar in ArabicContextForms) advances[scalar] = 200;
        return ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(advances);
    }

    private static IEnumerable<HtmlRenderText> ArabicContextText(HtmlRenderDocument document) =>
        document.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>();

    private static HtmlRenderShape ArabicContextBox(HtmlRenderDocument document) =>
        Assert.Single(document.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#probe");
}
