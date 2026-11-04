using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingNativeTextContextTests {
    private const string Family = "Native context fixture";
    private static OfficeRasterCanvas Metrics() => new(new OfficeRasterImage(1, 1), fonts:
        new OfficeFontFaceCollection().Add(Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120)));
    private static OfficeRichTextRun Run(string text, double threshold, OfficeColor color) =>
        new OfficeRichTextRun(text, 20, color, fontFamily: Family).WithAutomaticPairKerningThreshold(threshold);

    [Theory]
    [InlineData(false, 8, 20)]
    [InlineData(true, 8, 17.6)]
    [InlineData(true, 0, 20)]
    public void Native_font_context_shapes_across_color_and_threshold_boundaries_without_changing_annotations(
        bool native, double firstThreshold, double expectedWidth) {
        var paragraph = new OfficeRichTextParagraph(new[] { Run("A", firstThreshold, OfficeColor.Red),
            Run("V", firstThreshold == 0 ? 8 : 0, OfficeColor.Blue) }) { UsesPublisherKerningContext = native };
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 100, 100, Metrics().CreateDrawingTextMetrics());
        var line = Assert.Single(layout.Lines);
        Assert.Equal(expectedWidth, line.Width, 8);
        Assert.Equal(2, line.SourceTextEnd);
        Assert.Equal(new[] { OfficeColor.Red, OfficeColor.Blue }, line.Segments.Select(segment => segment.Color));
        Assert.Equal(paragraph.Runs.Select(run => run.AutomaticPairKerningThreshold),
            line.Segments.Select(segment => segment.AutomaticPairKerningThreshold));
        if (native) Assert.All(line.Segments, segment => {
            Assert.NotNull(segment.PreparedShaping);
            Assert.Equal(segment.Width, segment.PreparedShaping!.Measure(segment.FontSize), 8);
        });
        else Assert.All(line.Segments, segment => Assert.Null(segment.PreparedShaping));
    }

    [Fact]
    public void Word_across_source_styles_breaks_by_context_and_restarts_kerning_on_each_visual_line() {
        var paragraph = new OfficeRichTextParagraph(new[] { Run("A", 8, OfficeColor.Red), Run("VAV", 0, OfficeColor.Blue) }) {
            UsesPublisherKerningContext = true
        };
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 18, 200, Metrics().CreateDrawingTextMetrics());
        Assert.Equal(new[] { "AV", "A", "V" }, layout.Lines.Select(line => string.Concat(line.Segments.Select(segment => segment.Text))));
        Assert.Equal(new[] { 2, 3, 4 }, layout.Lines.Select(line => line.SourceTextEnd));
        Assert.Equal(new[] { 17.6, 10, 10 }, layout.Lines.Select(line => Math.Round(line.Width, 6)));
        Assert.Equal(OfficeColor.Red, layout.Lines[0].Segments[0].Color);
        Assert.Equal(OfficeColor.Blue, layout.Lines[0].Segments[1].Color);
        Assert.All(layout.Lines.Skip(1).SelectMany(line => line.Segments), segment => Assert.Equal(OfficeColor.Blue, segment.Color));
    }
    [Fact]
    public void Native_glyph_slices_keep_color_and_ink_positions_in_raster_and_svg() {
        var drawing = new OfficeDrawing(100, 80).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { Run("A", 8, OfficeColor.Red), Run("V", 0, OfficeColor.Blue) }) {
                UsesPublisherKerningContext = true
            }
        }, 10, 10, 80, 50);
        drawing.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120));
        var raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(12, 20));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(18, 20));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(24, 20));
        Assert.Equal(0, raster.GetPixel(27, 20).A);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing, 1D, OfficeSvgSizeUnit.Pixel);
        var document = System.Xml.Linq.XDocument.Parse(svg);
        Assert.Equal(new[] { "A", "V" }, document.Descendants()
            .Where(element => element.Attribute("aria-label") != null).Select(element => (string?)element.Attribute("aria-label")));
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(svg), out var restored));
        var svgRaster = OfficeDrawingRasterRenderer.Render(restored!);
        Assert.Equal(OfficeColor.Red, svgRaster.GetPixel(12, 20));
        Assert.Equal(OfficeColor.Blue, svgRaster.GetPixel(18, 20));
        Assert.Equal(OfficeColor.Blue, svgRaster.GetPixel(24, 20));
        Assert.Equal(0, svgRaster.GetPixel(27, 20).A);
    }

    [Theory]
    [InlineData(2D)]
    [InlineData(1.5D)]
    public void Native_raster_context_is_independent_of_export_resolution(double scale) {
        var drawing = new OfficeDrawing(100, 80).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { Run("A", 8, OfficeColor.Red), Run("V", 0, OfficeColor.Blue) }) {
                UsesPublisherKerningContext = true
            }
        }, 10, 10, 80, 50);
        drawing.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120));
        var raster = OfficeDrawingRasterRenderer.Render(drawing, scale);
        Assert.Equal(OfficeColor.Blue, raster.GetPixel((int)(18 * scale), (int)(20 * scale)));
        Assert.Equal(0, raster.GetPixel((int)(27 * scale), (int)(20 * scale)).A);
    }

    [Theory]
    [InlineData(OfficeTextTabAlignment.Right, 40D)]
    [InlineData(OfficeTextTabAlignment.Center, 70D)]
    [InlineData(OfficeTextTabAlignment.Character, 70D)]
    public void Native_aligned_tab_fields_keep_multipart_word_sizes(OfficeTextTabAlignment alignment, double start) {
        var paragraph = new OfficeRichTextParagraph(new[] { Run("\tV", 0, OfficeColor.Red),
            new OfficeRichTextRun("VA", 40, OfficeColor.Blue, fontFamily: Family).WithAutomaticPairKerningThreshold(0),
            Run("V", 0, OfficeColor.Black) }) { UsesPublisherKerningContext = true }
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(100, alignment, "A") }));
        var line = Assert.Single(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 200, 100,
            Metrics().CreateDrawingTextMetrics()).Lines);
        Assert.Equal(start, line.Segments.TakeWhile(segment => segment.Text.Length == 0).Sum(segment => segment.Width), 8);
        Assert.Equal(60, line.Segments.Where(segment => segment.Text.Length != 0).Sum(segment => segment.Width), 8);
        Assert.Equal(5, line.SourceTextEnd);
    }

    [Fact]
    public void Native_line_height_resolution_preserves_glyph_tokens_for_raw_svg_justification() {
        var fonts = new OfficeFontFaceCollection().Add(Family,
            ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120, includeSpace: true));
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), fonts: fonts);
        var layout = OfficeTextLayoutEngine.LayoutRichTextBlockWithMetrics(new[] { Run("AV AV AV", 8, OfficeColor.Red) },
            50, 100, 1.2, canvas.CreateDrawingTextMetrics().RichTextMeasurement.ForPublisherContext(), wrap: true);
        Assert.Equal(2, layout.Lines.Count);
        var output = new System.Text.StringBuilder().AppendSvgRichTextBlock(layout, 0, 0, 50, 100,
            horizontalAlignment: OfficeTextAlignment.Justify);
        var svg = System.Xml.Linq.XDocument.Parse("<svg>" + output + "</svg>");
        Assert.Equal("AVAVAV", string.Concat(svg.Descendants().Select(element => (string?)element.Attribute("aria-label"))).Replace(" ", ""));
    }

    [Fact]
    public void Native_character_tab_uses_the_painted_advance_into_its_delimiter() {
        var paragraph = new OfficeRichTextParagraph(new[] { Run("\tA", 8, OfficeColor.Red), Run("V", 0, OfficeColor.Blue) }) {
            UsesPublisherKerningContext = true
        }.WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(100, OfficeTextTabAlignment.Character, "V") }));
        var line = Assert.Single(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 200, 100,
            Metrics().CreateDrawingTextMetrics()).Lines);
        Assert.Equal(92.4, line.Segments.TakeWhile(segment => segment.Text.Length == 0).Sum(segment => segment.Width), 8);
        Assert.Equal(100, line.Segments.TakeWhile(segment => segment.Text != "V").Sum(segment => segment.Width), 8);
        Assert.Equal(OfficeColor.Blue, line.Segments.Single(segment => segment.Text == "V").Color);
    }

}
