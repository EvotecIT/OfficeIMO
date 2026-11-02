using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("line-height:2", 20D)]
    [InlineData("line-height:18px", 18D)]
    [InlineData("line-height:150%", 15D)]
    [InlineData("font:10px/1.8 Pinned", 18D)]
    public void HtmlPageMarginTypography_AuthoredLeadingTakesPrecedenceOverTheCallerDefault(
        string declaration, double expectedLineHeight) {
        HtmlRenderOptions options = PageMarginTypographyOptions();
        options.DefaultLineHeight = 2.5D;
        string html = PageMarginTypographySource("font:10px Pinned;" + declaration);

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options).Pages);
        HtmlRenderText margin = PageMarginTypographyText(page);

        Assert.Equal(10D, margin.Font.Size, 6);
        Assert.Equal(expectedLineHeight, margin.LineHeight, 6);
        Assert.True(margin.Height >= margin.LineHeight);
    }

    [Theory]
    [InlineData(null, 19D)]
    [InlineData(1.2D, 19.1875D)]
    [InlineData(2D, 32D)]
    public void HtmlPageMarginTypography_NormalUsesTheSelectedFaceUnlessTheCallerOverridesIt(
        double? defaultLineHeight, double expectedLineHeight) {
        HtmlRenderOptions options = PageMarginTypographyOptions();
        if (defaultLineHeight.HasValue) options.DefaultLineHeight = defaultLineHeight.Value;
        string html = PageMarginTypographySource("font:16px Pinned;line-height:normal", "Hgyp");

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options).Pages);
        HtmlRenderText margin = PageMarginTypographyText(page);

        // The same pinned face supplies the 19px normal box in NaturalLineHeight.
        Assert.Equal(expectedLineHeight, margin.LineHeight, 6);
        Assert.True(margin.Height >= margin.LineHeight);
    }

    [Fact]
    public void HtmlPageMarginTypography_LaterLonghandOverridesTheFontShorthandLeading() {
        string html = PageMarginTypographySource("font:10px/1.1 Pinned;line-height:18px");

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);

        Assert.Equal(18D, PageMarginTypographyText(page).LineHeight, 6);
    }

    [Fact]
    public void HtmlPageMarginTypography_FontShorthandResetsEarlierLeadingToNormal() {
        string html = PageMarginTypographySource("line-height:40px;font:16px Pinned", "Hgyp");

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);
        HtmlRenderText margin = PageMarginTypographyText(page);

        Assert.Equal(16D, margin.Font.Size, 6);
        Assert.Equal(19D, margin.LineHeight, 6);
    }

    [Fact]
    public void HtmlPageMarginTypography_ImportantLeadingSurvivesFirstPageAndLeftPageCascade() {
        const string html = """
            <style>
              @page :left { @bottom-center { line-height:150% !important; } }
              @page { size:300px 200px; margin:50px;
                @bottom-center { content:'Margin'; font:10px/1.1 Pinned; line-height:20px !important; }
              }
              @page :first { @bottom-center { font-size:12px; line-height:30px; } }
              body,p { margin:0; font:16px/20px Pinned; }
            </style>
            <p>First</p><p style='break-before:page'>Second</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions());

        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderText first = PageMarginTypographyText(rendered.Pages[0]);
        HtmlRenderText left = PageMarginTypographyText(rendered.Pages[1]);
        Assert.Equal(12D, first.Font.Size, 6);
        Assert.Equal(20D, first.LineHeight, 6);
        Assert.Equal(10D, left.Font.Size, 6);
        Assert.Equal(15D, left.LineHeight, 6);
    }

    [Fact]
    public void HtmlPageMarginTypography_TightWrappedFooterPaintsEveryWordAsAnArtifact() {
        const string footer = "ALPHA BETA GAMMA DELTA";
        string html = "<style>@page{size:130px 160px;margin:20px 20px 32px;"
            + "@bottom-left{content:'" + footer + "';font:10px/8px Pinned}"
            + "@bottom-center{content:'C';font:10px/8px Pinned}"
            + "@bottom-right{content:'R';font:10px/8px Pinned}}"
            + "body,p{margin:0;font:16px/20px Pinned}</style><p>BodyOnly</p>";
        var options = new HtmlToPdfOptions(PageMarginTypographyOptions()) {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText margin = PageMarginTypographyText(page, "bottom-left");

        Assert.Equal(8D, margin.LineHeight, 6);
        // Short leading advances the lines without shrinking their glyph-paint frames.
        Assert.True(margin.Height > 4D * margin.LineHeight);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Source == "@page @bottom-left"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        HtmlRenderSemanticGroup artifact = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @bottom-left");
        Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, artifact.Role);
        Assert.Equal("BodyOnly", rendered.Text.Trim());

        // The final glyph frame crosses the physical page and uses a clip group.
        // Walk that public drawing too; its children have local coordinates.
        string[] words = footer.Split(' ');
        OfficeDrawingText[] painted = PageMarginTypographyDrawingText(page.CreateDrawing())
            .Where(text => words.Contains(text.Text)).ToArray();
        Assert.Equal(PageMarginTypographyCompact(footer),
            PageMarginTypographyCompact(string.Concat(painted.Select(text => text.Text))));

        byte[] pdf = (document.ToPdf(options)).ToBytes();
        Assert.Equal("BodyOnly", PdfCore.PdfReadDocument.Open(pdf).ExtractText().Trim());
        var pdfPage = Assert.Single(PdfCore.PdfReadDocument.Open(pdf,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true }).Pages);
        string paintedFooter = string.Concat(pdfPage.GetTextSpans()
            .Where(span => span.X >= margin.X * 0.75D - 0.001D
                && span.X < (margin.X + margin.Width) * 0.75D - 0.001D
                && span.Y <= (page.Height - margin.Y) * 0.75D + 0.001D)
            .Select(span => span.Text));
        Assert.Equal(PageMarginTypographyCompact(footer), PageMarginTypographyCompact(paintedFooter));
    }

    [Fact]
    public void HtmlPageMarginTypography_CenteredTextUsesTheSamePositionAsOrdinaryText() {
        const string html = """
            <style>@page{size:300px 200px;margin:50px;
              @bottom-center{content:'Hgyp';font:16px/28px Pinned}}
              body,p{margin:0;font:16px/28px Pinned;text-align:center}
            </style><p>Hgyp</p>
            """;
        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);
        OfficeDrawingText[] text = page.CreateDrawing().Elements.OfType<OfficeDrawingText>().ToArray();
        Assert.Equal(2, text.Length);
        Assert.Equal(text[0].X, text[1].X, 6);
    }

    [Fact]
    public void HtmlPageMarginTypography_ZeroLeadingRetainsPositiveGlyphFrames() {
        string html = PageMarginTypographySource("font:10px/0 Pinned", "Hgyp");
        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);
        HtmlRenderText margin = PageMarginTypographyText(page);
        Assert.Equal(0D, margin.LineHeight);
        OfficeDrawingText text = Assert.Single(page.CreateDrawing().Elements.OfType<OfficeDrawingText>(),
            item => item.Text == "Hgyp");
        Assert.True(text.Height > 0D);
        Assert.Equal("Hgyp", text.Text);
    }

    [Fact]
    public void HtmlPageMarginTypography_CharacterUnitsDoNotInheritTheLastBodyFont() {
        double Measure(string body) {
            string html = "<style>@page{size:300px 200px;margin:50px;"
                + "@bottom-center{content:'Hgyp';font:16px/2ch Pinned}}"
                + "body,p{margin:0;font:16px/20px Pinned}</style>" + body;
            HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
                HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);
            return PageMarginTypographyText(page).LineHeight;
        }
        Assert.Equal(Measure("<p>Body</p>"), Measure(
            "<div style='container-type:size;width:100px;height:50px'>"
            + "<p style='font:32px/20px Pinned;margin-left:1ch'>Body</p></div>"), 6);
    }

    [Fact]
    public void HtmlPageMarginTypography_ZeroLeadingStillReportsTheSharedLineLimit() {
        string footer = string.Concat(Enumerable.Repeat("A ", 4100)) + "FINAL";
        string html = "<style>@page{size:28px 60px;margin:10px;@bottom-center{content:'"
            + footer + "';font:10px/0 Pinned}}body,p{margin:0}</style><p>B</p>";
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html),
            PageMarginTypographyOptions());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @bottom-center"
            && diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    private static HtmlRenderOptions PageMarginTypographyOptions() {
        HtmlRenderOptions options = NaturalLineOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.HonorCssPageRules = true;
        return options;
    }

    private static string PageMarginTypographySource(string declarations, string content = "Margin") =>
        "<style>@page{size:300px 200px;margin:50px;@bottom-center{content:'" + content
        + "';" + declarations + "}}body,p{margin:0;font:16px/20px Pinned}</style><p>BodyOnly</p>";

    private static HtmlRenderText PageMarginTypographyText(HtmlRenderPage page, string position = "bottom-center") =>
        Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(),
            text => text.Source == "@page @" + position);

    private static string PageMarginTypographyCompact(string text) =>
        string.Concat(text.Where(character => !char.IsWhiteSpace(character)));

    private static IEnumerable<OfficeDrawingText> PageMarginTypographyDrawingText(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingText text) yield return text;
            OfficeDrawing? nested = element is OfficeDrawingGroup group ? group.Drawing
                : element is OfficeDrawingEffectGroup effect ? effect.Drawing : null;
            if (nested != null) {
                foreach (OfficeDrawingText child in PageMarginTypographyDrawingText(nested)) yield return child;
            }
        }
    }
}
