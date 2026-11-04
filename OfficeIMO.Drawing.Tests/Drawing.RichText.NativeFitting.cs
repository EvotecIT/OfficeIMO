using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingNativeTextFittingTests {
    private const string Family = "Native fitting fixture";
    private static OfficeFontFaceCollection Fonts() => new OfficeFontFaceCollection().Add(Family,
        ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -200, includeSpace: true, reciprocal: true));
    private static OfficeRasterCanvas Metrics() => new(new OfficeRasterImage(1, 1), fonts: Fonts());
    private static OfficeRichTextParagraph Paragraph(string word, bool native = true) => new(
        new[] { new OfficeRichTextRun(word.Substring(0, 1), 20, OfficeColor.Red, fontFamily: Family).WithAutomaticPairKerningThreshold(8),
            new OfficeRichTextRun(word.Substring(1), 20, OfficeColor.Blue, fontFamily: Family).WithAutomaticPairKerningThreshold(0) }) {
        UsesPublisherKerningContext = native
    };
    private static string[] Lines(OfficeRichTextBlockLayout layout) => layout.Lines.Select(
        line => string.Concat(line.Segments.Select(segment => segment.Text))).ToArray();

    [Theory]
    [InlineData("AVAV", true, "AV|A|V")]
    [InlineData("AVAA", true, "AV|A|A")]
    [InlineData("AV", true, "A|V")]
    [InlineData("AVV", true, "A|V|V")]
    [InlineData("AVAV", false, "A|V|A|V")]
    public void Native_broken_words_fit_with_lookahead_but_paint_only_the_visible_line(string word, bool native, string expected) {
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { Paragraph(word, native) },
            12.5, 200, Metrics().CreateDrawingTextMetrics());
        Assert.Equal(expected.Split('|'), Lines(layout));
        Assert.Equal(word.Length, layout.Lines.Last().SourceTextEnd);
        if (expected.StartsWith("AV|", StringComparison.Ordinal)) {
            var first = layout.Lines[0];
            // AV fits at 12 with its following VA pair, but visible AV paints
            // at 16. The last glyph must lose the pair against the next line.
            Assert.Equal(16, first.Width, 8);
            Assert.Equal(6, first.Segments[0].Width, 8);
            Assert.Equal(10, first.Segments[1].Width, 8);
            Assert.Equal(2, first.SourceTextEnd);
        }
    }

    [Fact]
    public void Native_region_continuation_keeps_selected_lines_and_exact_source_positions() {
        var metrics = Metrics().CreateDrawingTextMetrics();
        var flow = new OfficeRichTextFlow(new[] { Paragraph("AVAV") }, _ => { });
        var first = Assert.Single(flow.Take(12.5, 24, metrics, CancellationToken.None));
        Assert.Equal("AV", string.Concat(first.Runs.Select(run => run.Text)));
        Assert.Equal(2, flow.CharacterPosition);
        var painted = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { first }, 12.5, 24, metrics);
        Assert.Equal(new[] { "AV" }, Lines(painted));
        Assert.Equal(16, Assert.Single(painted.Lines).Width, 8);
        Assert.False(painted.Clipped);
        Assert.True(flow.HasRemaining);
        foreach (var expected in new[] { "A", "V" }) {
            var next = Assert.Single(flow.Take(12.5, 24, metrics, CancellationToken.None));
            Assert.Equal(new[] { expected }, Lines(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { next }, 12.5, 24, metrics)));
        }
        Assert.Equal(4, flow.CharacterPosition);
        Assert.False(flow.HasRemaining);
    }

    [Fact]
    public void Native_region_slices_do_not_move_a_shortened_word_onto_the_preceding_line() {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("A ", 8, OfficeColor.Red, fontFamily: Family).WithAutomaticPairKerningThreshold(8),
            new OfficeRichTextRun("AVAV", 20, OfficeColor.Blue, fontFamily: Family).WithAutomaticPairKerningThreshold(0)
        }) { UsesPublisherKerningContext = true };
        var metrics = Metrics().CreateDrawingTextMetrics();
        var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => { });
        var first = Assert.Single(flow.Take(29, 34, metrics, CancellationToken.None));
        Assert.Equal("A AV", string.Concat(first.Runs.Select(run => run.Text)));
        Assert.Equal(4, flow.CharacterPosition);
        Assert.Equal(new[] { "A ", "AV" }, Lines(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { first }, 29, 34, metrics)));
        var next = Assert.Single(flow.Take(29, 34, metrics, CancellationToken.None));
        Assert.Equal("AV", string.Concat(next.Runs.Select(run => run.Text)));
        Assert.Equal(6, flow.CharacterPosition);
        Assert.False(flow.HasRemaining);
    }

    [Fact]
    public void Native_selected_line_does_not_repaint_discarded_wrap_whitespace() {
        var metrics = Metrics().CreateDrawingTextMetrics();
        var flow = new OfficeRichTextFlow(new[] { Paragraph("A AVAV").WithTabStops(new OfficeTextTabStops(Array.Empty<OfficeTextTabStop>())) }, _ => { });
        var first = Assert.Single(flow.Take(12.5, 24, metrics, CancellationToken.None));
        Assert.Equal("A ", string.Concat(first.Runs.Select(run => run.Text)));
        Assert.Equal(2, flow.CharacterPosition);
        var line = Assert.Single(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { first }, 12.5, 24, metrics).Lines);
        Assert.Equal("A", string.Concat(line.Segments.Select(segment => segment.Text)));
        Assert.Equal(10, line.Width, 8);
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Center, .25)]
    [InlineData(OfficeTextAlignment.Right, .5)]
    public void Native_selected_line_alignment_uses_its_admission_advance(OfficeTextAlignment alignment, double expectedOffset) {
        var source = Paragraph("AVAV");
        var paragraph = new OfficeRichTextParagraph(source.Runs, alignment) { UsesPublisherKerningContext = true };
        var metrics = Metrics().CreateDrawingTextMetrics();
        var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => { });
        var first = Assert.Single(flow.Take(12.5, 24, metrics, CancellationToken.None));
        Assert.Equal(2, flow.CharacterPosition);
        var line = Assert.Single(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { first }, 12.5, 24, metrics).Lines);
        Assert.Equal("AV", string.Concat(line.Segments.Select(segment => segment.Text)));
        Assert.Equal(16, line.Width, 8);
        Assert.Equal(expectedOffset, line.OffsetX, 8);
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(1.5D)]
    [InlineData(2D)]
    public void Native_selected_lines_keep_visible_overhang_in_raster_and_svg(double scale) {
        var selected = new OfficeRichTextFlow(new[] { Paragraph("AVAV") }, _ => { })
            .Take(12.5, 24, Metrics().CreateDrawingTextMetrics(), CancellationToken.None);
        var drawing = new OfficeDrawing(100, 80).AddRichTextParagraphs(selected, 10, 10, 12.5, 24);
        drawing.Fonts.AddRange(Fonts());
        var raster = OfficeDrawingRasterRenderer.Render(drawing, scale);
        Assert.Equal(OfficeColor.Blue, raster.GetPixel((int)(23 * scale), (int)(20 * scale)));
        Assert.Equal(0, raster.GetPixel((int)(25 * scale), (int)(20 * scale)).A);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing, scale, OfficeSvgSizeUnit.Pixel);
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(svg), out var restored));
        var svgRaster = OfficeDrawingRasterRenderer.Render(restored!);
        Assert.Equal(OfficeColor.Blue, svgRaster.GetPixel((int)(23 * scale), (int)(20 * scale)));
        Assert.Equal(0, svgRaster.GetPixel((int)(25 * scale), (int)(20 * scale)).A);
    }

    [Fact]
    public void Native_selected_line_overhang_survives_an_isolated_effect_surface() {
        var selected = new OfficeRichTextFlow(new[] { Paragraph("AVAV") }, _ => { })
            .Take(12.5, 24, Metrics().CreateDrawingTextMetrics(), CancellationToken.None);
        var inner = new OfficeDrawing(12.5, 24).AddRichTextParagraphs(selected, 0, 0, 12.5, 24);
        ((OfficeDrawingRichText)inner.Elements[0]).NativePublisherFrame = true;
        var drawing = new OfficeDrawing(100, 80).AddEffectDrawing(inner, OfficeTransform.Translate(10, 10), opacity: .5);
        drawing.Fonts.AddRange(Fonts());
        var raster = OfficeDrawingRasterRenderer.Render(drawing);
        var pixel = raster.GetPixel(23, 20);
        Assert.Equal(OfficeColor.Blue.R, pixel.R);
        Assert.Equal(OfficeColor.Blue.G, pixel.G);
        Assert.Equal(OfficeColor.Blue.B, pixel.B);
        Assert.InRange(pixel.A, 127, 128);
    }

    [Theory]
    [InlineData(OfficeTextTabAlignment.Right, 0D)]
    [InlineData(OfficeTextTabAlignment.Center, 15D)]
    [InlineData(OfficeTextTabAlignment.Character, 0D)]
    public void Native_selected_wraps_preserve_aligned_tab_fields_and_leader_advances(OfficeTextTabAlignment alignment, double expectedAdvance) {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("\tAAAA AAVA", 20, OfficeColor.Black, fontFamily: Family).WithAutomaticPairKerningThreshold(0)
        }) { UsesPublisherKerningContext = true }.WithTabStops(new OfficeTextTabStops(new[] {
            new OfficeTextTabStop(60, alignment, "V").WithLeader("V")
        }));
        var metrics = Metrics().CreateDrawingTextMetrics();
        var selected = Assert.Single(new OfficeRichTextFlow(new[] { paragraph }, _ => { })
            .Take(70, 100, metrics, CancellationToken.None));
        Assert.Equal("\tAAAA AAVA", string.Concat(selected.Runs.Select(run => run.Text)));
        var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { selected }, 70, 100, metrics);
        Assert.Equal(2, layout.Lines.Count);
        var first = layout.Lines[0];
        var body = Assert.Single(first.Segments, segment => segment.Text == "AAAA ");
        Assert.Equal(expectedAdvance, first.Segments.TakeWhile(segment => !ReferenceEquals(segment, body)).Sum(segment => segment.Width), 8);
        Assert.Equal(50 + expectedAdvance, first.Width, 8);
        Assert.Equal(6, first.SourceTextEnd);
    }

    [Fact]
    public void Reframing_selected_native_text_recomputes_alignment_for_the_new_width() {
        var source = Paragraph("AVAV");
        var paragraph = new OfficeRichTextParagraph(source.Runs, OfficeTextAlignment.Center) { UsesPublisherKerningContext = true };
        var metrics = Metrics().CreateDrawingTextMetrics();
        var selected = Assert.Single(new OfficeRichTextFlow(new[] { paragraph }, _ => { }).Take(12.5, 24, metrics, default));
        var line = Assert.Single(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { selected }, 20, 24, metrics).Lines);
        Assert.Equal(16, line.Width, 8);
        Assert.Equal(2, line.OffsetX, 8);
    }

    [Fact]
    public void Replacing_selected_native_tab_settings_recomputes_tab_placement() {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("\tAAAA AAVA", 20, OfficeColor.Black, fontFamily: Family).WithAutomaticPairKerningThreshold(0)
        }) { UsesPublisherKerningContext = true }.WithTabStops(new OfficeTextTabStops(new[] {
            new OfficeTextTabStop(60, OfficeTextTabAlignment.Right)
        }));
        var metrics = Metrics().CreateDrawingTextMetrics();
        var selected = Assert.Single(new OfficeRichTextFlow(new[] { paragraph }, _ => { }).Take(70, 100, metrics, default))
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(20) }));
        var line = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { selected }, 70, 100, metrics).Lines[0];
        var body = Assert.Single(line.Segments, segment => segment.Text == "AAAA ");
        Assert.Equal(20, line.Segments.TakeWhile(segment => !ReferenceEquals(segment, body)).Sum(segment => segment.Width), 8);
    }

    [Fact]
    public void Reusing_selected_native_paragraphs_honors_a_public_no_wrap_request() {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("A A A", 20, OfficeColor.Black, fontFamily: Family).WithAutomaticPairKerningThreshold(0)
        }) { UsesPublisherKerningContext = true };
        var metrics = Metrics().CreateDrawingTextMetrics();
        var selected = new OfficeRichTextFlow(new[] { paragraph }, _ => { }).Take(30, 60, metrics, default);
        var drawing = new OfficeDrawing(100, 80).AddRichTextParagraphs(selected, 0, 0, 30, 60, wrapText: false);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        var line = Assert.Single(OfficeDrawingTextLayout.CreateWithMetrics(text, 30, 60, metrics).Lines);
        Assert.Equal("A A A", string.Concat(line.Segments.Select(segment => segment.Text)));
        Assert.Equal(50, line.Width, 8);
        Assert.Equal(5, line.SourceTextEnd);
    }
}
