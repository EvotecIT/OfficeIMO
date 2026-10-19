using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string PreparedArabicColorWord = "<span style='color:red'>م</span><span style='color:blue'>رحبا</span>";

    [Fact]
    public void HtmlArabicProvider_ColorBoundariesUseWholeContextAdvancesAndLogicalSource() {
        byte[] font = PreparedArabicProviderFont();
        var provider = new PreparedArabicContextProvider(font);
        const string content = "<a href='https://example.com/arabic' style='text-decoration:none'>"
            + PreparedArabicColorWord + "</a>";
        HtmlRenderDocument colored = RenderArabicContext(content, provider: provider, fontData: font);
        HtmlRenderDocument plain = RenderArabicContext(ArabicContextWord,
            provider: new PreparedArabicContextProvider(font), fontData: font);
        HtmlRenderText[] text = ArabicContextText(colored).ToArray();
        double expected = 20D * 1000D / provider.Program.UnitsPerEm;

        Assert.Equal(expected, ArabicContextText(plain).Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        Assert.Equal(expected, text.Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        HtmlRenderText meem = Assert.Single(text, item => item.Color == OfficeColor.Red);
        HtmlRenderText rest = Assert.Single(text, item => item.Color == OfficeColor.Blue);
        Assert.Equal(rest.X + rest.TextAdvanceWidth!.Value, meem.X, 6);
        Assert.Equal(ArabicContextWord, colored.Text.Trim());
        Assert.Equal(plain.Text, colored.Text);
        Assert.All(text, item => Assert.Equal("https://example.com/arabic", item.LinkUri));
        Assert.Contains(provider.Requests, request => request.Text.Contains(ArabicContextWord, StringComparison.Ordinal));
        colored.RequireNoLoss();
    }

    [Theory]
    [InlineData("25px", false)]
    [InlineData("min-content", false)]
    [InlineData("max-content", false)]
    [InlineData("max-content", true)]
    public void HtmlArabicProvider_ContextWidthsDriveFitIntrinsicAndFirstLine(string width, bool firstLine) {
        byte[] font = PreparedArabicProviderFont();
        string css = firstLine ? "#probe::first-line{color:green}" : string.Empty;
        var provider = new PreparedArabicContextProvider(font);
        HtmlRenderDocument colored = RenderArabicContext("Z " + PreparedArabicColorWord, width, css, provider, font);
        HtmlRenderDocument plain = RenderArabicContext("Z " + ArabicContextWord, width, css,
            new PreparedArabicContextProvider(font), font);
        HtmlRenderText[] text = ArabicContextText(colored).ToArray();
        double fullAdvance = 20D * 1200D / provider.Program.UnitsPerEm;
        double wordAdvance = 20D * 1000D / provider.Program.UnitsPerEm;

        Assert.Equal(ArabicContextText(plain).Sum(item => item.TextAdvanceWidth ?? 0D),
            text.Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        Assert.Equal(ArabicContextBox(plain).Width, ArabicContextBox(colored).Width, 6);
        if (width == "min-content") {
            Assert.Equal(wordAdvance, ArabicContextBox(colored).Width, 6);
            Assert.Equal(wordAdvance, text.Where(item => item.Color == OfficeColor.Red || item.Color == OfficeColor.Blue)
                .Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        } else Assert.Equal(fullAdvance, text.Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        if (width == "max-content") Assert.Equal(fullAdvance, ArabicContextBox(colored).Width, 6);
        // The whole contextual word fits after Z and its space; isolated run advances do not.
        if (width == "25px") Assert.Single(text.Select(item => item.Y).Distinct());
        if (firstLine) Assert.All(text, item => Assert.Equal(OfficeColor.Green, item.Color));
        Assert.Equal(plain.Text, colored.Text);
        if (width == "25px") Assert.Equal("Z " + ArabicContextWord, colored.Text.Trim());
        colored.RequireNoLoss();
    }

    [Theory]
    [InlineData(" ")]
    [InlineData("\u200C")]
    [InlineData("<br>")]
    [InlineData("<span style='display:inline-block;width:8px;height:8px'></span>")]
    public void HtmlArabicProvider_AuthoredBoundariesKeepIndependentContext(string separator) {
        byte[] font = PreparedArabicProviderFont();
        var provider = new PreparedArabicContextProvider(font);
        HtmlRenderDocument rendered = RenderArabicContext("<span style='color:red'>م</span>" + separator
            + "<span style='color:blue'>رحبا</span>", provider: provider, fontData: font);
        HtmlRenderText[] text = ArabicContextText(rendered).ToArray();
        double isolatedMeemAdvance = 20D * 700D / provider.Program.UnitsPerEm;

        Assert.DoesNotContain(provider.Requests, request => request.Text.Contains(ArabicContextWord, StringComparison.Ordinal));
        Assert.Equal(isolatedMeemAdvance,
            text.Where(item => item.Color == OfficeColor.Red).Sum(item => item.TextAdvanceWidth ?? 0D), 6);
        Assert.Equal(ArabicContextWord, OfficeArabicTextShaper.ToLogicalText(rendered.Text)
            .Replace(" ", string.Empty).Replace("\u200C", string.Empty)
            .Replace("\r", string.Empty).Replace("\n", string.Empty).Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlArabicProvider_BackendsReceiveUnmirroredLogicalBracketSource() {
        byte[] font = PreparedArabicProviderFont();
        var provider = new PreparedArabicContextProvider(font);
        const string logical = ArabicContextWord + " (";
        HtmlRenderDocument rendered = RenderArabicContext(logical, provider: provider, fontData: font);
        foreach (HtmlRenderPage page in rendered.Pages) {
            OfficeDrawing drawing = page.CreateDrawing().ApplyImageExportOptions(new OfficeImageExportOptions {
                TextShapingProvider = provider
            });
            OfficeDrawingRasterRenderer.Render(drawing);
        }

        Assert.Equal(logical, rendered.Text.TrimEnd('\r', '\n'));
        Assert.Contains(provider.Requests, request => request.Text.Contains('('));
        // This catches the observed second mirroring at the backend, while allowing a
        // retained result to avoid requesting the provider again during painting.
        Assert.DoesNotContain(provider.Requests, request => request.Text.Contains(')'));
    }

    private static byte[] PreparedArabicProviderFont() {
        var advances = new Dictionary<int, ushort>();
        foreach (char scalar in ArabicContextWord + ArabicContextForms + "\uFEE1\uFEAD\uFEA1\uFE8F\uFE8D\u200C")
            advances[scalar] = 700;
        foreach (char scalar in " Z()123") advances[scalar] = 100;
        return ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(advances);
    }

    // The fixture is a real decoded TrueType program. This deliberately small provider
    // makes context retention observable without claiming native shaping correctness.
    private sealed class PreparedArabicContextProvider : IOfficeTextShapingProvider {
        private readonly bool _mergeFirstCluster;
        internal PreparedArabicContextProvider(byte[] fontData, bool mergeFirstCluster = false) {
            Program = OfficeTrueTypeFont.TryLoad(fontData)!;
            _mergeFirstCluster = mergeFirstCluster;
        }
        internal IOfficeFontProgram Program { get; }
        internal List<OfficeTextShapingRequest> Requests { get; } = new();
        internal int GlyphId(char scalar) {
            if (!Program.TryGetGlyphMetrics(scalar, out int glyphId, out _))
                throw new InvalidOperationException("The declared synthetic font lacks scalar " + ((int)scalar).ToString("X4"));
            return glyphId;
        }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            request.CancellationToken.ThrowIfCancellationRequested();
            Requests.Add(request);
            string text = request.Text;
            if (text.Length == 0) return null;
            int context = text.IndexOf(ArabicContextWord, StringComparison.Ordinal);
            OfficeTextDirection direction = request.Direction == OfficeTextDirection.Auto
                ? OfficeTextElements.ResolveBaseDirection(text) : request.Direction;
            var glyphs = new List<OfficeShapedGlyph>();
            for (int index = 0; index < text.Length; index++) {
                char scalar = text[index];
                if (OfficeTextElements.ContainsBidiControl(scalar.ToString())) continue;
                int position = context < 0 ? -1 : index - context;
                bool contextual = position >= 0 && position < ArabicContextWord.Length;
                if (contextual && _mergeFirstCluster && position == 1) continue;
                char paint = contextual ? ArabicContextForms[position] : scalar;
                if (direction == OfficeTextDirection.RightToLeft)
                    paint = OfficeBidiTextResolver.MirrorText(paint.ToString())[0];
                int advance = contextual ? 160 + position * 20
                    : scalar == ' ' || scalar == 'Z' || scalar == '(' || scalar == ')' ? 100 : 700;
                string unicode = contextual && _mergeFirstCluster && position == 0
                    ? text.Substring(index, 2) : scalar.ToString();
                if (contextual && _mergeFirstCluster && position == 0) advance += 180;
                glyphs.Add(new OfficeShapedGlyph(GlyphId(paint), unicode, index, advance,
                    offsetX: contextual ? 11 + position : 0, offsetY: contextual ? -7 - position : 0));
            }
            if (direction == OfficeTextDirection.RightToLeft) glyphs.Reverse();
            return new OfficeTextShapingResult(glyphs, direction);
        }
    }
}
