using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("سلام 123.45")]
    [InlineData("سلام123.45")]
    public void HtmlManagedTypography_NaturalBidiPartitionsTheMeasuredFontPens(string source) {
        byte[] data = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf"));
        HtmlRenderDocument rendered = RenderTypographyPartition("<bdi lang='ar' dir='rtl'>" + source + "</bdi>", data);
        HtmlRenderText[] text = TypographyPartitionText(rendered);
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(data)!;
        OfficeTextShapingResult shaped = Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest(source, "Typography", data, false, font.UnitsPerEm,
                OfficeTextDirection.RightToLeft, "ar")));
        OfficeTrueTypeFont.ShapedTextRun expected = font.CreateShapedTextRun(source, shaped);
        PaintedTestGlyph[] painted = ObservePaintedTrueTypeGlyphs(rendered, text);

        rendered.RequireNoLoss();
        Assert.Equal(source, rendered.Text.Trim());
        Assert.Equal(expected.Measure(16D), text.Sum(item => item.TextAdvanceWidth), 8);
        Assert.Equal(expected._glyphs.Select(glyph => (int)glyph.GlyphId), painted.Select(glyph => glyph.GlyphId));
        var expectedOrigins = new List<double>();
        double pen = 0D;
        foreach (OfficeTrueTypeFont.PositionedGlyph glyph in expected._glyphs) {
            expectedOrigins.Add(pen + glyph.OffsetX * 16D / font.UnitsPerEm);
            pen += glyph.AdvanceWidth * 16D / font.UnitsPerEm;
        }
        double expectedLeft = expectedOrigins.Min();
        double actualLeft = painted.Min(glyph => glyph.X);
        for (int index = 0; index < painted.Length; index++) {
            Assert.Equal(expectedOrigins[index] - expectedLeft, painted[index].X - actualLeft, 8);
        }
        Assert.All(text, item => Assert.Equal("Typography", item.Font.FamilyName));
    }

    [Fact]
    public void HtmlManagedTypography_CombiningTokenDoesNotReshapeIndependentlyMeasuredNeighbours() {
        const string marked = "Cafe\u0301";
        const string tail = " delivery office — 24 items.";
        byte[] data = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf"));
        HtmlRenderDocument combined = RenderTypographyPartition(marked + tail, data);
        HtmlRenderDocument markedOnly = RenderTypographyPartition(marked, data);
        HtmlRenderDocument tailOnly = RenderTypographyPartition("<span style='white-space:pre'>" + tail + "</span>", data);
        HtmlRenderText[] combinedText = TypographyPartitionText(combined);
        double expectedAdvance = TypographyPartitionText(markedOnly).Sum(item => item.TextAdvanceWidth)
            + TypographyPartitionText(tailOnly).Sum(item => item.TextAdvanceWidth);

        combined.RequireNoLoss();
        markedOnly.RequireNoLoss();
        tailOnly.RequireNoLoss();
        Assert.Equal(marked + tail, combined.Text.Trim());
        Assert.Equal(expectedAdvance, combinedText.Sum(item => item.TextAdvanceWidth), 8);
        Assert.Contains(combinedText, item => item.Text.Contains("\u0301") && item.PreparedShaping != null);
        Assert.All(combinedText, item => Assert.Equal("Typography", item.Font.FamilyName));
    }

    private static HtmlRenderDocument RenderTypographyPartition(string content, byte[] fontData) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:480px 180px;margin:0}html,body,p{margin:0;padding:0}"
            + "p{font:16px/24px Typography}</style><p dir='ltr'>" + content + "</p>");
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        options.Fonts.Add("Typography", fontData);
        return HtmlRenderEngine.Render(document, options);
    }

    private static HtmlRenderText[] TypographyPartitionText(HtmlRenderDocument rendered) => rendered.Pages
        .SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>().ToArray();
}
