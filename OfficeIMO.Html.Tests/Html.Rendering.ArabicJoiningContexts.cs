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

        Assert.Equal(ArabicContextGlyphIds("\uFEE3"), ArabicContextPaintedGlyphIds(new[] { meem }));
        Assert.Equal(ArabicContextGlyphIds("\uFE8E\uFE92\uFEA3\uFEAE"), ArabicContextPaintedGlyphIds(remainder));
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
        Assert.Equal(ArabicContextGlyphIds(new string(ArabicContextForms.Reverse().ToArray())),
            ArabicContextPaintedGlyphIds(ArabicContextText(rendered)));
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
        Assert.Equal(ArabicContextGlyphIds("\uFEE1"),
            ArabicContextPaintedGlyphIds(text.Where(item => item.Color == OfficeColor.Red)));
        Assert.Equal(ArabicContextGlyphIds("\uFE8E\uFE92\uFEA3\uFEAD"),
            ArabicContextPaintedGlyphIds(text.Where(item => item.Color == OfficeColor.Blue)));
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ArabicContextFont())!;
        Assert.Equal(font.Measure("\uFEE1", 20D),
            text.Where(item => item.Color == OfficeColor.Red).Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        Assert.Equal(font.Measure("\uFEAD\uFEA3\uFE92\uFE8E", 20D),
            text.Where(item => item.Color == OfficeColor.Blue).Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        Assert.Equal(ArabicContextWord, OfficeArabicTextShaper.ToLogicalText(rendered.Text)
            .Replace(" ", string.Empty).Replace("\u200C", string.Empty).Replace("\r", string.Empty).Replace("\n", string.Empty).Trim());
        Assert.True(!rendered.HasLoss, string.Join(" | ", rendered.Diagnostics.Select(item => item.Source + ":" + item.Detail)));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlArabicJoining_TransparentJoinerAndMarkRemainInLogicalSource(bool followingWord) {
        const string logical = "م\u064E\u200Dرحبا";
        string suffix = followingWord ? " " + ArabicContextWord : string.Empty;
        HtmlRenderDocument rendered = RenderArabicContext("<span style='color:red'>م\u064E\u200D</span>"
            + "<span style='color:blue'>رحبا" + suffix + "</span>");
        HtmlRenderText[] text = ArabicContextText(rendered).ToArray();
        int[] painted = ArabicContextPaintedGlyphIds(text.Where(item => item.PreparedShaping == null));
        Assert.Contains(Assert.Single(ArabicContextGlyphIds("\uFEE3")), painted);
        Assert.Contains(Assert.Single(ArabicContextGlyphIds("\uFEAE")), painted);
        // The declined complete context retains both styled pieces. A later
        // word in the same source run can still keep its accepted glyphs.
        if (followingWord) Assert.True(text.Any(item => item.PreparedShaping?.Text.Trim() == ArabicContextWord),
            string.Join(" | ", text.Select(item => "Paint=" + item.Text + "; Prepared=" + (item.PreparedShaping?.Text ?? "<none>"))));
        Assert.Equal(logical + suffix, rendered.Text.Trim());
        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    private static HtmlRenderDocument RenderArabicContext(string content, string width = "300px", string extraCss = "",
        IOfficeTextShapingProvider? provider = null, byte[]? fontData = null) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 400D, ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false,
            DefaultFontFamily = ManagedTextShapingTestAssets.FamilyName, TextShapingProvider = provider
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, fontData ?? ArabicContextFont());
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
        // These are ordinary positive joining fixtures. Whitespace and joiner
        // controls must be nonpainting while their source and advances remain.
        return ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(advances, new[] { 32, 0x200C, 0x200D });
    }

    // Inspect the glyphs actually carried into painting. Retained provider results
    // keep logical scene text; the scalar fallback keeps presentation-form text.
    private static int[] ArabicContextPaintedGlyphIds(IEnumerable<HtmlRenderText> text) =>
        text.OrderBy(item => item.X).SelectMany(item => item.PreparedShaping != null
            ? item.PreparedShaping.Result.Glyphs.Select(glyph => glyph.GlyphId)
            : ArabicContextGlyphIds(item.Text)).ToArray();

    private static int[] ArabicContextGlyphIds(string forms) {
        IOfficeFontProgram font = OfficeTrueTypeFont.TryLoad(ArabicContextFont())!;
        return forms.Select(scalar => {
            Assert.True(font.TryGetGlyphMetrics(scalar, out int glyph, out _));
            return glyph;
        }).ToArray();
    }

    private static IEnumerable<HtmlRenderText> ArabicContextText(HtmlRenderDocument document) =>
        document.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>();

    private static HtmlRenderShape ArabicContextBox(HtmlRenderDocument document) =>
        Assert.Single(document.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#probe");
}
