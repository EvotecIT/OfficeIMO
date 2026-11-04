using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextKerningTests {
    private const string Family = "Run kerning fixture";
    private static OfficeFontFaceCollection Fonts(bool fallback = false) {
        var fonts = new OfficeFontFaceCollection();
        if (fallback) fonts.Add(Family, ManagedTextShapingTestAssets.CreateFont('A'));
        fonts.Add(fallback ? "Run fallback" : Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120));
        return fonts;
    }
    private static OfficeRichTextRun Run(string text, double size, double? threshold) =>
        new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: Family).WithAutomaticPairKerningThreshold(threshold);
    private static OfficeRasterCanvas Canvas(OfficeFontFaceCollection? fonts = null, OfficeRasterImage? image = null) =>
        new OfficeRasterCanvas(image ?? new OfficeRasterImage(100, 80), fonts: fonts ?? Fonts());

    [Theory]
    [InlineData("unspecified", true)]
    [InlineData("zero", false)]
    [InlineData("below-size", true)]
    [InlineData("equal-size", true)]
    [InlineData("above-size", false)]
    public void Explicit_threshold_uses_inclusive_size_boundary_and_scales_with_drawing_units(string policy, bool enabled) {
        double? threshold = policy switch { "zero" => 0, "below-size" => 8, "equal-size" => 12, "above-size" => 14, _ => null };
        var paragraph = new OfficeRichTextParagraph(new[] { Run("AV", 12, threshold) });
        var canvas = Canvas();
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 100, 100, canvas.CreateDrawingTextMetrics());
        var segment = Assert.Single(Assert.Single(layout.Lines).Segments);
        // The fixture has 500-unit advances, a -120-unit AV pair and 1000 units/em.
        Assert.Equal(enabled ? 10.56 : 12, segment.Width, 8);
        Assert.Equal(threshold, segment.AutomaticPairKerningThreshold);
        var scaled = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 200, 200,
            canvas.CreateDrawingTextMetrics(), scale: 2);
        var scaledSegment = Assert.Single(Assert.Single(scaled.Lines).Segments);
        Assert.Equal(segment.Width * 2, scaledSegment.Width, 8);
        Assert.Equal(threshold * 2, scaledSegment.AutomaticPairKerningThreshold);
    }

    [Fact]
    public void Mixed_thresholds_keep_distinct_shaping_contexts_and_source_endpoints_when_wrapping() {
        var canvas = Canvas();
        var paragraph = new OfficeRichTextParagraph(new[] { Run("AV", 12, 0), Run("AV", 12, 8) });
        var mixed = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 23, 100, canvas.CreateDrawingTextMetrics());
        var line = Assert.Single(mixed.Lines);
        Assert.Equal(22.56, line.Width, 8); Assert.Equal(4, line.SourceTextEnd);
        Assert.Equal(new double?[] { 0, 8 }, line.Segments.Select(segment => segment.AutomaticPairKerningThreshold));
        var off = new OfficeRichTextParagraph(paragraph.Runs.Select(run => run.WithAutomaticPairKerningThreshold(0)).ToArray());
        var wrapped = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { off }, 23, 100, canvas.CreateDrawingTextMetrics());
        Assert.Equal(2, wrapped.Lines.Count);
        Assert.Equal(new[] { 2, 4 }, wrapped.Lines.Select(value => value.SourceTextEnd));
        Assert.Equal("AVAV", string.Concat(wrapped.Lines.SelectMany(value => value.Segments).Select(value => value.Text)));
    }

    [Fact]
    public void Raster_glyph_positions_and_cached_ink_follow_each_segment_policy() {
        var canvas = Canvas();
        foreach (var policy in new[] { (Threshold: 0D, Right: 18D), (Threshold: 8D, Right: 15.6D), (Threshold: 0D, Right: 18D) }) {
            var paragraph = new OfficeRichTextParagraph(new[] { Run("AV", 20, policy.Threshold) });
            var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 80, 50, canvas.CreateDrawingTextMetrics());
            var ink = Assert.Single(OfficeDrawingTextLayout.PlacedRasterParagraphPaintBounds(layout, canvas, 0, 0, 50, OfficeTextVerticalAlignment.Top));
            Assert.Equal(policy.Right, ink.Right, 8);
            var image = new OfficeRasterImage(100, 80);
            OfficeTextBlockRenderer.DrawRasterRichTextBlock(Canvas(image: image), layout, 10, 10, 80, 50);
            int right = Enumerable.Range(0, image.Width).Where(x => Enumerable.Range(0, image.Height).Any(y => image.GetPixel(x, y).A > 0)).Max();
            Assert.InRange(right, (int)Math.Floor(10 + policy.Right) - 1, (int)Math.Ceiling(10 + policy.Right) - 1);
        }
    }

    [Fact]
    public void Fallback_faces_do_not_kern_across_font_boundaries() {
        var fonts = Fonts(fallback: true);
        var canvas = Canvas(fonts);
        OfficeRichTextRun FallbackRun(double size) => new OfficeRichTextRun("AV", size, OfficeColor.Black,
            fontFamily: Family + ", Run fallback").WithAutomaticPairKerningThreshold(8);
        var paragraph = new OfficeRichTextParagraph(new[] { FallbackRun(12) });
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 100, 100, canvas.CreateDrawingTextMetrics());
        Assert.Equal(12, Assert.Single(layout.Lines).Width, 8);
        Assert.Equal("AV", string.Concat(layout.Lines.SelectMany(line => line.Segments).Select(segment => segment.Text)));
        var image = new OfficeRasterImage(100, 80);
        var large = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] {
            new OfficeRichTextParagraph(new[] { FallbackRun(20) })
        }, 100, 100, canvas.CreateDrawingTextMetrics());
        OfficeTextBlockRenderer.DrawRasterRichTextBlock(Canvas(fonts, image), large, 10, 10, 80, 50);
        Assert.Equal(OfficeColor.Black, image.GetPixel(26, 20));
    }

    [Theory]
    [InlineData(17)]
    [InlineData(18)]
    [InlineData(19)]
    public void Justification_uses_the_advances_of_the_tokens_that_are_actually_painted(int width) {
        var fonts = new OfficeFontFaceCollection().Add(Family,
            ManagedTextShapingTestAssets.CreateFontWithKerning('A', ' ', -120));
        var paragraph = new OfficeRichTextParagraph(new[] { Run("A A A", 12, 8) }, OfficeTextAlignment.Justify);
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, width, 100,
            Canvas(fonts).CreateDrawingTextMetrics());
        Assert.Equal(2, layout.Lines.Count);
        Assert.Equal(width, layout.Lines[0].Width, 8);
        Assert.Equal(4, layout.Lines[0].SourceTextEnd);
        Assert.Equal(5, layout.Lines[1].SourceTextEnd);
        Assert.Equal("A A", string.Concat(layout.Lines[0].Segments.Select(segment => segment.Text)));
    }

    [Fact]
    public void Shrink_to_fit_considers_the_higher_fitting_interval_after_kerning_activates() {
        var drawing = new OfficeDrawing(100, 100).AddRichText(new[] { Run("AVAVAV", 18, 14) },
            0, 0, 38, 80, wrapText: false, shrinkToFit: true);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        var layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, 38, 80, Canvas());
        var segment = Assert.Single(Assert.Single(layout.Lines).Segments);
        Assert.InRange(segment.FontSize, 14.38, 14.40);
        // Frame fitting retains the shared 0.01-unit width allowance.
        Assert.InRange(layout.Width, 37.95, 38.01);
        Assert.Equal(14D, segment.AutomaticPairKerningThreshold);
    }

    [Fact]
    public void Svg_emits_run_font_features_as_css_that_browsers_apply() {
        var drawing = new OfficeDrawing(100, 100)
            .AddRichText(new[] { Run("AV", 20, 0) }, 10, 10, 80, 30)
            .AddRichText(new[] { Run("AV", 20, 8) }, 10, 50, 80, 30);
        drawing.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120));
        XDocument svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        var text = svg.Descendants().Where(element => element.Name.LocalName == "text" && element.Value == "AV").ToArray();
        Assert.Equal(2, text.Length);
        Assert.Equal("font-feature-settings:\"kern\" 0", text[0].Attribute("style")?.Value);
        Assert.Equal("font-feature-settings:\"kern\" 1", text[1].Attribute("style")?.Value);
    }
}
