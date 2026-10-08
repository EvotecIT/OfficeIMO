using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("3px", "8px", 3D, 8D)]
    [InlineData(".125em", "25%", 4D, 8D)]
    [InlineData("0", "-4px", 1D, -4D)]
    public void TextDecorationLengthsPaintOnceAtTheAlphabeticOffset(string thickness, string offset, double expectedThickness, double expectedOffset) {
        string html = "<a style='font:32px/40px Arial;text-decoration:underline " + thickness
            + " solid red;text-underline-offset:" + offset + ";text-decoration-skip-ink:none' href='https://example.test'>Geometry</a>";
        HtmlRenderDocument scene = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 600 });
        var visuals = scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).ToArray();
        HtmlRenderText text = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Geometry");
        HtmlRenderShape line = Assert.Single(visuals.OfType<HtmlRenderShape>(), t => t.Source?.EndsWith(":decoration:underline", StringComparison.Ordinal) == true);
        Assert.Equal(expectedThickness, line.Height, 6);
        Assert.Equal(text.Y + 32D + expectedOffset, line.Y, 6);
        Assert.Equal(text.TextAdvanceWidth!.Value, line.Width, 6);
        Assert.Equal(OfficeColor.Red, line.Shape.FillColor);
        Assert.False(text.Font.IsUnderline);
        Assert.DoesNotContain(scene.Diagnostics, d => d.Code == "HtmlRenderTextDecorationThicknessApproximated");
        PdfReadDocument pdf = PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { ViewportWidth = 600 }));
        Assert.Single(pdf.Pages.SelectMany(p => p.GetLinkAnnotations()));
        Assert.Contains("Geometry", string.Concat(pdf.Pages.Select(p => p.ExtractText())));
    }

    [Fact]
    public void NestedDecorationPropagatesColorAndGeometryThroughRelativePaintAndWrappedLines() {
        const string html = "<div style='width:85px;font:20px/24px Arial'><span style='position:relative;left:7px;top:5px;text-decoration:overline underline 3px solid red;text-underline-offset:6px;text-decoration-skip-ink:none'><span style='color:blue'>Alpha beta gamma delta</span></span></div>";
        HtmlRenderDocument scene = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 200 });
        var visuals = scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).ToArray();
        HtmlRenderText[] texts = visuals.OfType<HtmlRenderText>().Where(t => !string.IsNullOrWhiteSpace(t.Text)).ToArray();
        Assert.True(texts.Select(t => t.Y).Distinct().Count() > 1);
        HtmlRenderShape[] lines = visuals.OfType<HtmlRenderShape>().Where(t => t.Source?.Contains(":decoration:", StringComparison.Ordinal) == true).ToArray();
        Assert.Equal(texts.Length * 2, lines.Length);
        Assert.All(lines, line => {
            Assert.Equal(3D, line.Height);
            Assert.Equal(OfficeColor.Red, line.Shape.FillColor);
            double offset = line.Source!.EndsWith(":overline", StringComparison.Ordinal) ? -1.5D : 26D;
            Assert.Contains(texts, text => Math.Abs(text.X - line.X) < 0.000001D
                && Math.Abs(text.Y + offset - line.Y) < 0.000001D);
        });
        Assert.All(texts, text => Assert.Equal(OfficeColor.Blue, text.Color));
    }

    [Fact]
    public void CombinedSolidDecorationPreservesTextAndPaintStacking() {
        var scene = HtmlRenderTestDriver.Render("<span style='font:32px/40px Arial;text-decoration:underline overline line-through 3px solid red;text-underline-offset:8px;text-decoration-skip-ink:none'>Layer</span>", new HtmlRenderOptions());
        var visuals = scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).ToArray();
        HtmlRenderText text = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Layer");
        HtmlRenderShape[] lines = visuals.OfType<HtmlRenderShape>().Where(s => s.Source?.Contains(":decoration:", StringComparison.Ordinal) == true).ToArray();
        Assert.Equal(3, lines.Length);
        foreach (HtmlRenderShape line in lines) {
            bool strike = line.Source!.EndsWith(":line-through", StringComparison.Ordinal);
            Assert.Equal(strike, Array.IndexOf(visuals, line) > Array.IndexOf(visuals, text));
            if (strike) Assert.Equal(text.Y + 32D - 9.6D - 1.5D, line.Y, 6);
        }
        Assert.DoesNotContain(scene.Diagnostics, d => d.Code is "HtmlRenderTextDecorationThicknessApproximated" or "HtmlRenderTextDecorationLineUnsupported");
    }

    [Theory]
    [InlineData("solid")]
    [InlineData("double")]
    [InlineData("dashed")]
    [InlineData("dotted")]
    [InlineData("wavy")]
    public void TextDecorationPatternsExportToPdfWithoutDuplicateLinksOrText(string pattern) {
        string html = "<a href='https://example.test' style='font:32px/40px Arial;text-decoration:underline 3px "
            + pattern + " red;text-underline-offset:8px;text-decoration-skip-ink:none'>Pattern</a>";
        PdfReadDocument pdf = PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { ViewportWidth = 600 }));
        Assert.Single(pdf.Pages.SelectMany(p => p.GetLinkAnnotations()));
        Assert.Equal("Pattern", string.Concat(pdf.Pages.Select(p => p.ExtractText())).Trim());
    }

    [Fact]
    public void TextDecorationBidiControlsRetainAutomaticPaintAndReportTheSpecialization() {
        var scene = HtmlRenderTestDriver.Render("<span style='text-decoration:underline 3px solid red;text-underline-offset:8px'>\u202Eabc\u202C</span>", new HtmlRenderOptions());
        Assert.Contains(scene.Diagnostics, d => d.Code == "HtmlRenderTextDecorationThicknessApproximated"
            && d.Detail?.Contains("bidi", StringComparison.Ordinal) == true);
        Assert.DoesNotContain(scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).OfType<HtmlRenderShape>(),
            s => s.Source?.Contains(":decoration:", StringComparison.Ordinal) == true);
    }

    [Theory]
    [InlineData("text-decoration:underline from-font", "from-font")]
    [InlineData("text-decoration-line:underline;text-decoration-thickness:FROM-FONT", "from-font")]
    [InlineData("text-decoration:overline 3px;writing-mode:vertical-rl", "vertical")]
    [InlineData("text-decoration:underline 3px;text-decoration-skip-ink:all", "skipping")]
    public void SpecializedDecorationGeometryRetainsActionableLoss(string declarations, string detail) {
        var scene = HtmlRenderTestDriver.Render("<span style='" + declarations + "'>Boundary</span>", new HtmlRenderOptions());
        Assert.Contains(scene.Diagnostics, d => d.Code == "HtmlRenderTextDecorationThicknessApproximated" && d.Detail?.Contains(detail, StringComparison.Ordinal) == true);
        Assert.True(scene.HasLoss);
    }
}
