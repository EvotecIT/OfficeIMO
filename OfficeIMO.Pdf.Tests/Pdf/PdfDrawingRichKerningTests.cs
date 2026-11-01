using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingRichKerningTests {
    private const string Family = "Rich kerning fixture";
    private static OfficeRichTextRun Run(string text, double size, double? threshold, string family = Family) =>
        new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: family).WithAutomaticPairKerningThreshold(threshold);
    private static byte[] Write(OfficeDrawing drawing) => PdfDocument.Create(new PdfOptions {
        PageWidth = 100, PageHeight = 80, MarginLeft = 0, MarginRight = 0, MarginTop = 0, MarginBottom = 0
    }.UseRenderingProfile(new OfficeRenderingProfile("Rich kerning", drawing.Fonts)))
        .Canvas(canvas => canvas.Drawing(drawing, 0, 0, 100, 80)).ToBytes();

    [Theory]
    [InlineData("unspecified", true)]
    [InlineData("zero", false)]
    [InlineData("below-size", true)]
    [InlineData("equal-size", true)]
    [InlineData("above-size", false)]
    public void Rich_threshold_controls_painted_glyph_positions_alongside_measurement(string policy, bool enabled) {
        double? threshold = policy switch { "zero" => 0, "below-size" => 8, "equal-size" => 20, "above-size" => 24, _ => null };
        var drawing = new OfficeDrawing(100, 80).AddRichText(new[] { Run("AV", 20, threshold) }, 10, 10, 80, 50);
        drawing.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120));
        byte[] bytes = Write(drawing);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("AV", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(10, letters[0].StartBaseLine.X, 3);
        // 500 units/em advance; enabled AV adjustment is -120 at 1000 units/em.
        Assert.Equal(enabled ? 17.6 : 20, letters[1].StartBaseLine.X, 3);
        Assert.Equal(letters[0].StartBaseLine.Y, letters[1].StartBaseLine.Y, 3);
    }

    [Fact]
    public void Mixed_run_policies_keep_one_line_and_independent_pdf_glyph_advances() {
        var drawing = new OfficeDrawing(100, 80).AddRichText(new[] { Run("AV", 12, 0), Run("AV", 12, 8) },
            10, 10, 23, 50);
        drawing.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(Write(drawing));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("AVAV", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => Assert.Equal(letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        Assert.Equal(new[] { 10D, 16D, 22D, 26.56D }, letters.Select(letter => System.Math.Round(letter.StartBaseLine.X, 3)));
    }

    [Fact]
    public void Rich_policy_preserves_selected_fallback_faces_without_cross_font_kerning() {
        var drawing = new OfficeDrawing(100, 80).AddRichText(new[] { Run("AV", 20, 8, Family + ", Rich fallback") },
            10, 10, 80, 50);
        drawing.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.Fonts.Add("Rich fallback", ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120));
        byte[] bytes = Write(drawing);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("AV", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(20, letters[1].StartBaseLine.X, 3);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(bytes).Pages[0].ToDrawing(), background: OfficeColor.White);
        Assert.Equal(OfficeColor.Black, image.GetPixel(25, 20));
    }
    [Fact]
    public void Unwrapped_paragraph_keeps_disabled_kerning_italic_overhang_outside_frame() {
        var run = new OfficeRichTextRun("AV", 20, OfficeColor.Black, italic: true, fontFamily: Family)
            .WithAutomaticPairKerningThreshold(0);
        var drawing = new OfficeDrawing(100, 80).AddRichTextParagraphs(
            new[] { new OfficeRichTextParagraph(new[] { run }) }, 10, 10, 5, 50, wrapText: false);
        drawing.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -240));
        byte[] bytes = Write(drawing);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("AV", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(20, letters[1].StartBaseLine.X, 3);
        // V starts at20pt. Its8pt rectangle and14pt height, under PDF's
        // synthetic1/3 oblique, put the top-right ink at32 2/3pt.
        string content = PdfTestContentText.ReadDecoded(bytes);
        var clips = System.Text.RegularExpressions.Regex.Matches(content,
            @"(-?[0-9.]+) (-?[0-9.]+) (-?[0-9.]+) (-?[0-9.]+) re W n")
            .Cast<System.Text.RegularExpressions.Match>()
            .Select(match => (Left: double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture),
                Width: double.Parse(match.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture))).ToArray();
        Assert.Contains(clips, clip => clip.Left <= 10D + .001D
            && clip.Left + clip.Width >= 32D + 2D / 3D - .001D);
    }

    [Theory]
    [InlineData(false, 8, 20)]
    [InlineData(true, 8, 17.6)]
    [InlineData(true, 0, 20)]
    public void Native_context_retains_embedded_searchable_glyphs_across_paint_boundaries(bool native, double firstThreshold, double secondX) {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("A", 20, OfficeColor.Red, fontFamily: Family).WithAutomaticPairKerningThreshold(firstThreshold),
            new OfficeRichTextRun("V", 20, OfficeColor.Blue, fontFamily: Family).WithAutomaticPairKerningThreshold(firstThreshold == 0 ? 8 : 0)
        }) { UsesPublisherKerningContext = native };
        var drawing = new OfficeDrawing(100, 80).AddRichTextParagraphs(new[] { paragraph }, 10, 10, 80, 50);
        drawing.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(Write(drawing));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("AV", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(2, letters.Count);
        Assert.Equal(10, letters[0].StartBaseLine.X, 3);
        Assert.Equal(secondX, letters[1].StartBaseLine.X, 3);
        Assert.Equal((1D, 0D, 0D), letters[0].Color.ToRGBValues());
        Assert.Equal((0D, 0D, 1D), letters[1].Color.ToRGBValues());
        Assert.Equal(letters[0].FontName, letters[1].FontName);
    }

}
