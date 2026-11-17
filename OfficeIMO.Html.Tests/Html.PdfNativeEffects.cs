using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

/// <summary>Effects the PDF writer represents natively stay native in HTML-to-PDF output.</summary>
public sealed class HtmlPdfNativeEffectsTests {
    private const string GradientRasterized = "HtmlPdfTranslucentGradientRasterized";
    private const string DrawingEffectRasterized = "HtmlPdfDrawingEffectRasterized";

    [Theory]
    [InlineData("linear-gradient(to right,red,transparent)", 150, 100)]
    [InlineData("radial-gradient(circle closest-side,red,transparent)", 175, 100)]
    public void TranslucentGradientsStayVectorShadingWithoutDarkFringes(string gradient, int x, int y) {
        string html = "<style>@page{size:200px 200px;margin:0}body{margin:0}</style>"
            + "<div style='width:200px;height:200px;background:" + gradient + "'></div>";

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(html).ConvertToPdfWithEvidence(new HtmlPdfExportOptions());
        byte[] pdf = result.ToBytes();
        OfficeRasterImage page = OfficeDrawingRasterRenderer.Render(PdfCore.PdfPageImageRenderer.RenderPage(pdf), background: OfficeColor.White);

        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == GradientRasterized);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        string syntax = Encoding.GetEncoding(28591).GetString(pdf);
        Assert.Contains("/Shading", syntax, StringComparison.Ordinal);
        Assert.Contains("/S /Luminosity", syntax, StringComparison.Ordinal);
        // CSS interpolates in premultiplied alpha: a quarter-opaque red over white stays
        // pink. Interpolating toward the keyword's black would darken the red channel.
        OfficeColor sample = Sample(page, 200D, x, y);
        Assert.True(sample.R >= 245, sample.ToString());
        Assert.InRange(sample.G, (byte)150, (byte)230);
        Assert.InRange(sample.G - sample.B, -4, 4);
    }

    [Theory]
    [InlineData(null, 1875)]
    [InlineData(150D, 938)]
    public void GradientsBeyondNativeShadingRasterizeAtTheExportDensity(double? targetDpi, int expectedWidth) {
        // 300 hard-edged periods materialize more stops than one PDF shading function holds.
        const string html = "<style>@page{size:600px 8px;margin:0}body{margin:0}</style>"
            + "<div style='width:600px;height:4px;background:repeating-linear-gradient(90deg,rgba(255,0,0,.5) 0 1px,transparent 1px 2px)'></div>";
        var options = new HtmlPdfExportOptions { MaxGradientStops = 4096, TargetDpi = targetDpi };

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(html).ConvertToPdfWithEvidence(options);
        byte[] pdf = result.ToBytes();

        Assert.Contains(result.Report.Warnings, warning => warning.Code == GradientRasterized);
        PdfCore.PdfExtractedImage image = Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(pdf), item => !item.IsImageMask);
        Assert.InRange(image.Width, expectedWidth - 1, expectedWidth + 1);

        options.MaximumRasterPixels = 1000L;
        Assert.Throws<OfficeImageExportLimitException>(() => HtmlConversionDocument.Parse(html).ConvertToPdfWithEvidence(options));
    }

    [Fact]
    public void InlineSvgBlendModesCompositeWithThePageUnderneath() {
        PdfCore.PdfDocumentConversionResult result = ConvertBlendFixture(BlendSvg, inline: true, out OfficeRasterImage page, out byte[] pdf);

        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == DrawingEffectRasterized);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        Assert.Contains("/BM /Multiply", Encoding.GetEncoding(28591).GetString(pdf), StringComparison.Ordinal);
        // Cyan multiplied with the yellow page background is green.
        AssertColor(OfficeColor.FromRgb(0, 255, 0), Sample(page, 80D, 40, 40));
    }

    [Fact]
    public void SvgImageBlendModesStayInsideTheImage() {
        PdfCore.PdfDocumentConversionResult result = ConvertBlendFixture(BlendSvg, inline: false, out OfficeRasterImage page, out byte[] pdf);

        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == DrawingEffectRasterized);
        Assert.Contains("/BM /Multiply", Encoding.GetEncoding(28591).GetString(pdf), StringComparison.Ordinal);
        // An SVG image is its own surface, so its multiply has only a transparent backdrop.
        AssertColor(OfficeColor.FromRgb(0, 255, 255), Sample(page, 80D, 40, 40));
        AssertColor(OfficeColor.FromRgb(255, 255, 0), Sample(page, 80D, 75, 75));
    }

    [Theory]
    [InlineData("font-variant:small-caps", "smcp")]
    [InlineData("font-variant-caps:all-small-caps", "c2sc")]
    public void FontSmallCapsUseTheFaceGlyphsAndKeepTheSourceText(string caps, string feature) {
        string html = FontFace("SmallCapsFace", "SourceSansPro-Regular.otf", "otf")
            + "<p style='margin:0;font:24px SmallCapsFace;" + caps + "'>Annual Report</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderText[] texts = Texts(rendered).Where(item => item.Text.Trim().Length > 0).ToArray();
        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(html).ConvertToPdfWithEvidence(new HtmlPdfExportOptions());
        byte[] pdf = result.ToBytes();

        Assert.Equal("AnnualReport", string.Concat(texts.Select(item => item.Text)).Replace(" ", string.Empty));
        Assert.All(texts, text => Assert.True(text.FeatureSettings.TryGetValue(feature, out int value) && value == 1));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontVariantApproximated);
        Assert.Equal("Annual Report", PdfCore.PdfDocument.Load(new MemoryStream(pdf)).ExtractText().Trim());
    }

    [Fact]
    public void LetterSpacedFontSmallCapsPlaceEachGlyphAtItsSmallCapAdvance() {
        string html = FontFace("SmallCapsFace", "SourceSansPro-Regular.otf", "otf")
            + "<p style='margin:0;font:24px SmallCapsFace;font-variant:small-caps;letter-spacing:2px'>lll</p>"
            + "<p style='margin:0;font:24px SmallCapsFace;letter-spacing:2px'>lll</p>";

        HtmlRenderText[] glyphs = Texts(HtmlRenderTestDriver.Render(html)).Where(item => item.Text == "l").ToArray();

        Assert.Equal(6, glyphs.Length);
        double smallCapAdvance = glyphs[1].X - glyphs[0].X;
        double lowercaseAdvance = glyphs[4].X - glyphs[3].X;
        // A small-cap L is far wider than a lowercase l; spacing must follow the drawn glyph.
        Assert.True(smallCapAdvance > lowercaseAdvance * 1.5D, smallCapAdvance + " vs " + lowercaseAdvance);
    }

    [Fact]
    public void SynthesizedSmallCapsExtractTheSourceCase() {
        // Carlito has c2sc but no smcp, so lowercase letters are synthesized.
        string html = FontFace("NoSmallCaps", "Carlito-Regular.ttf", "ttf")
            + "<p style='margin:0;font:24px NoSmallCaps;font-variant:small-caps'>Annual Report</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderText[] texts = Texts(rendered).Where(item => item.Text.Trim().Length > 0).ToArray();
        byte[] pdf = HtmlConversionDocument.Parse(html).ConvertToPdfWithEvidence(new HtmlPdfExportOptions()).ToBytes();

        Assert.Equal("ANNUALREPORT", string.Concat(texts.Select(item => item.Text)).Replace(" ", string.Empty));
        Assert.True(Assert.Single(texts, item => item.Text == "NNUAL").Font.Size < Assert.Single(texts, item => item.Text == "A").Font.Size);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontVariantApproximated);
        Assert.Equal("Annual Report", PdfCore.PdfDocument.Load(new MemoryStream(pdf)).ExtractText().Trim());
    }

    private const string BlendSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='80' height='80' viewBox='0 0 80 80'>"
        + "<rect x='10' y='10' width='60' height='60' fill='#00ffff' style='mix-blend-mode:multiply'/></svg>";

    private static PdfCore.PdfDocumentConversionResult ConvertBlendFixture(string svg, bool inline,
        out OfficeRasterImage page, out byte[] pdf) {
        string image = inline
            ? svg.Replace("<svg ", "<svg style='display:block' ")
            : "<img style='display:block' width='80' height='80' src='data:image/svg+xml;base64,"
                + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + "'>";
        string html = "<style>@page{size:80px 80px;margin:0}body{margin:0;background:#ffff00}</style>" + image;
        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(html).ConvertToPdfWithEvidence(new HtmlPdfExportOptions());
        pdf = result.ToBytes();
        page = OfficeDrawingRasterRenderer.Render(PdfCore.PdfPageImageRenderer.RenderPage(pdf), background: OfficeColor.White);
        return result;
    }

    // The page raster's scale is independent of CSS pixels; sample by CSS coordinate.
    private static OfficeColor Sample(OfficeRasterImage page, double cssPageWidth, int x, int y) {
        double scale = page.Width / cssPageWidth;
        return page.GetPixel((int)Math.Floor(x * scale), (int)Math.Floor(y * scale));
    }

    private static void AssertColor(OfficeColor expected, OfficeColor actual) {
        Assert.True(Math.Abs(expected.R - actual.R) <= 6 && Math.Abs(expected.G - actual.G) <= 6 && Math.Abs(expected.B - actual.B) <= 6,
            "Expected " + expected + " but found " + actual);
    }

    private static string FontFace(string family, string file, string format) =>
        "<style>@font-face{font-family:" + family + ";src:url('data:font/" + format + ";base64,"
        + Convert.ToBase64String(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", file))) + "')}</style>";

    private static IEnumerable<HtmlRenderText> Texts(HtmlRenderDocument rendered) =>
        rendered.Pages.SelectMany(page => Flatten(page.Scene)).OfType<HtmlRenderText>();

    private static IEnumerable<HtmlRenderVisual> Flatten(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderVisual child in Flatten(children)) yield return child;
        }
    }
}
