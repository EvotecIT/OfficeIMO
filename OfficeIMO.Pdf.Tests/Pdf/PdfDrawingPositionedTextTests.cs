using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingPositionedTextTests {
    [Fact]
    public void CanvasDrawingKeepsPreparedGlyphInkAndLogicalTextThroughProfileAliases() {
        const string family = "Prepared Font";
        byte[] retainedData = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 'B');
        OfficeTrueTypeFont program = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(retainedData));
        var result = new OfficeTextShapingResult(new[] {
            // Logical A deliberately paints the wider B outline, followed by a second cluster glyph.
            new OfficeShapedGlyph(2, "A", 0, 1000, 100, 200),
            OfficeShapedGlyph.CreatePositionedUnicodeContinuation(1, 0, 0, 0, 200, -100)
        }, OfficeTextDirection.LeftToRight);
        var prepared = new OfficePreparedTextShaping(program, "A", result);
        var drawing = new OfficeDrawing(100D, 80D).AddPositionedText("A", 10D, 20D, 80D, 40D,
            new OfficeFontInfo(family, 20D), OfficeColor.Black, textAdvanceWidth: prepared.Measure(20D));
        drawing.Fonts.Add(family, retainedData);
        drawing.MarkLastTextWithPreparedShaping(prepared);

        // A scoped profile alias would clone with a conflicting program and clear the retained carrier.
        var profileFonts = new OfficeFontFaceCollection().Add(family,
            ManagedTextShapingTestAssets.CreateFont('A'), OfficeFontStyle.Regular,
            new OfficeFontUnicodeRangeSet(new[] { new OfficeFontUnicodeRange('A', 'A') }));
        Assert.NotEqual(family, Assert.Single(profileFonts.Faces).ResourceFamilyName);
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(100D), PageHeight = global::OfficeIMO.OfficeLength.Points(80D),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(0D), MarginRight = global::OfficeIMO.OfficeLength.Points(0D), MarginTop = global::OfficeIMO.OfficeLength.Points(0D), MarginBottom = global::OfficeIMO.OfficeLength.Points(0D)
        }.UseRenderingProfile(new OfficeRenderingProfile("prepared-alias", profileFonts));
        byte[] bytes = PdfDocument.Create(options)
            .Canvas(canvas => canvas.Drawing(drawing, global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(100D), global::OfficeIMO.OfficeLength.Points(80D))).ToBytes();
        PdfReadDocument loaded = PdfReadDocument.Open(bytes);
        Assert.Equal("A", loaded.ExtractText().Trim());
        Assert.Empty(PdfDocument.Load(bytes).Reader.Images());

        // Independent rectangle geometry proves the selected glyph IDs, offsets, pen advance and continuation.
        // Positioned text uses the alphabetic baseline at frame Y + source font size = 40.
        OfficeShape first = OfficeShape.Rectangle(12D, 14D);
        first.FillColor = OfficeColor.Black;
        first.StrokeColor = null;
        OfficeShape continuation = OfficeShape.Rectangle(8D, 14D);
        continuation.FillColor = OfficeColor.Black;
        continuation.StrokeColor = null;
        var reference = new OfficeDrawing(100D, 80D)
            .AddShape(first, 12D, 22D).AddShape(continuation, 34D, 28D);
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(reference, background: OfficeColor.White);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(
            loaded.Pages[0].ToDrawing(), background: OfficeColor.White);
        Assert.Equal(OfficeColor.Black, actual.GetPixel(15, 25));
        Assert.Equal(OfficeColor.Black, actual.GetPixel(38, 32));
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }

    [Fact]
    public void SinglePositionedLineKeepsItsExplicitAdvance() {
        var drawing = new OfficeDrawing(100, 80).AddPositionedText("A", 10, 10, 80, 60,
            new OfficeFontInfo("Helvetica", 20), alignment: OfficeTextAlignment.Center, textAdvanceWidth: 50);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(80), MarginLeft = global::OfficeIMO.OfficeLength.Points(0), MarginTop = global::OfficeIMO.OfficeLength.Points(0), MarginRight = global::OfficeIMO.OfficeLength.Points(0), MarginBottom = global::OfficeIMO.OfficeLength.Points(0)
        }).Compose(c => c.Page(p => p.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letter = Assert.Single(pdf.GetPage(1).Letters);
        Assert.InRange(letter.Width, 49.99D, 50.01D);
        Assert.InRange(letter.Location.X, 24.99D, 25.01D);
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Center, "\n")]
    [InlineData(OfficeTextAlignment.Right, "\r\n")]
    [InlineData(OfficeTextAlignment.Right, "\r")]
    public void PositionedHardLinesRetainTheirOwnAdvance(OfficeTextAlignment alignment, string separator) {
        var drawing = new OfficeDrawing(100, 80).AddPositionedText("AAAA" + separator + "A", 10, 10, 80, 60,
            new OfficeFontInfo("Helvetica", 20), alignment: alignment, lineHeight: 25);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(100), PageHeight = global::OfficeIMO.OfficeLength.Points(80), MarginLeft = global::OfficeIMO.OfficeLength.Points(0), MarginTop = global::OfficeIMO.OfficeLength.Points(0), MarginRight = global::OfficeIMO.OfficeLength.Points(0), MarginBottom = global::OfficeIMO.OfficeLength.Points(0)
        }).Compose(c => c.Page(p => p.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("AAAAA", string.Concat(letters.Select(letter => letter.Value)));
        Assert.InRange(Math.Abs(letters[0].Width - letters[4].Width), 0, .01D);
        Assert.InRange(letters[0].Width, 13.3D, 13.4D);
        double firstAdvance = letters[0].Width * 4;
        double lastAdvance = letters[4].Width;
        double factor = alignment == OfficeTextAlignment.Center ? .5D : 1D;
        Assert.InRange(Math.Abs(letters[0].Location.X - (10 + (80 - firstAdvance) * factor)), 0, .01D);
        Assert.InRange(Math.Abs(letters[4].Location.X - (10 + (80 - lastAdvance) * factor)), 0, .01D);
        Assert.InRange(Math.Abs(letters[0].Location.Y - letters[4].Location.Y - 25), 0, .01D);
    }
}
