using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("", 1000)]
    [InlineData("0 0 1 1 re W n 10 10 1 1 re W n ", 20)]
    [InlineData("0 0 1 1 re W n ", 20)]
    public void RenderPage_NonpaintingSpacedTextDoesNotConsumePositioningBudget(string clip, int x) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        string text = new string('A', PdfReadLimits.Default.MaxPositionedTextCharactersPerPage + 1);
        byte[] pdf = BuildSingleStreamPdf(clip + "BT /F1 10 Tf 1 Tc " + x + " 100 Td (" + text + ") Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = 1 }
        });
        OfficeDrawing drawing = document.Pages[0].ToDrawing();
        Assert.Empty(drawing.Elements);
    }

    [Theory]
    [InlineData(-70, 1)]
    [InlineData(300, -13)]
    public void RenderPage_PartiallyVisibleRunChargesOnlyProjectedGlyphs(int x, int spacing) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        string value = new string('A', 100);
        byte[] pdf = BuildSingleStreamPdf($"BT /F1 10 Tf {spacing} Tc {x} 100 Td ({value}) Tj ET", resources, font);
        const int visibleGlyphs = 35;
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = visibleGlyphs }
        });
        OfficeDrawing actual = document.Pages[0].ToDrawing();
        var reference = new StringBuilder();
        for (int index = 0; index < value.Length; index++) {
            int origin = x + index * (6 + spacing);
            reference.Append($"BT /F1 10 Tf {origin} 100 Td (A) Tj ET ");
        }
        OfficeDrawing expected = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(reference.ToString(), resources, font));
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected).GetPixels(), OfficeDrawingRasterRenderer.Render(actual).GetPixels());
        PdfReadDocument limited = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = visibleGlyphs - 1 }
        });
        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => limited.Pages[0].ToDrawing());
        Assert.Equal(PdfReadLimitKind.PositionedTextCharacters, error.Kind);
        Assert.Equal(visibleGlyphs, error.Actual);
    }

    [Fact]
    public void SpacedTextProjectionObservesInvocationCancellationWhenCallerOmitsToken() {
        using var cancellation = new System.Threading.CancellationTokenSource();
        PdfReadPage page = PdfReadDocument.Open(BuildSingleStreamPdf("", "<< >>")).Pages[0];
        var budget = new PdfReadPage.PageContentBudget(page, cancellation.Token);
        var span = new PdfTextSpan("AB", "F1", 10, 20, 100, 14, OfficeColor.Black, true, 0D, "Courier", null,
            characterAdvances: new[] { 7D, 7D }, canScaleAggregateAdvance: false,
            glyphCharacterLengths: new[] { 1, 1 }, glyphPaintedAdvances: new CancellingGlyphWidths(cancellation));
        var drawing = new OfficeDrawing(240, 200);
        var method = typeof(PdfReadPage).GetMethod("AddTextSpan", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
        var exception = Assert.Throws<System.Reflection.TargetInvocationException>(() => method.Invoke(null,
            new object[] { drawing, 200D, span, budget, default(System.Threading.CancellationToken) }));
        Assert.IsType<OperationCanceledException>(exception.InnerException);
        Assert.Empty(drawing.Elements);
    }

    private sealed class CancellingGlyphWidths : System.Collections.Generic.IReadOnlyList<double> {
        private readonly System.Threading.CancellationTokenSource _cancellation;
        internal CancellingGlyphWidths(System.Threading.CancellationTokenSource cancellation) => _cancellation = cancellation;
        public int Count => 2;
        public double this[int index] { get { _cancellation.Cancel(); return 6D; } }
        public System.Collections.Generic.IEnumerator<double> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [Theory]
    [InlineData("06280628", "بب")]
    [InlineData("0915093F", "कि")]
    [InlineData("00610301", "a\u0301")]
    public void RenderPage_SpacingRetainsContextualTextRuns(string unicode, string expectedText) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier /ToUnicode 6 0 R >>\nendobj";
        string cmap = "2 beginbfchar\n<41> <" + unicode.Substring(0, 4) + ">\n<42> <" + unicode.Substring(4) + ">\nendbfchar";
        byte[] pdf = BuildSingleStreamPdf("BT /F1 20 Tf 2 Tw 20 100 Td (AB) Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font, BuildStreamObject(6, "<<", cmap));
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = 1 }
        });
        PdfTextSpan span = Assert.Single(document.Pages[0].GetTextSpans());
        Assert.Equal(expectedText, span.Text);
        Assert.False(span.CanScaleAggregateAdvance);
        OfficeDrawing drawing = document.Pages[0].ToDrawing();
        // No word separator occurs: Tw must not change shaping, nor consume an expansion budget.
        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(expectedText, text.Text);
        var expected = new OfficeDrawing(drawing.Width, drawing.Height).AddText(text.Text, text.X, text.Y,
            text.Width, text.Height, text.Font, text.Color, wrapText: false);
        expected.Fonts.AddRange(drawing.Fonts);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected).GetPixels(), OfficeDrawingRasterRenderer.Render(drawing).GetPixels());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RenderPage_PositioningBudgetCountsPaintedExpansionOnly(bool actualText) {
        string font = actualText
            ? "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj"
            : "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier /FirstChar 65 /LastChar 66 /Widths [0 0] >>\nendobj";
        string content = actualText ? "/Span << /ActualText (Long replacement text) >> BDC " : "";
        content += "BT /F1 10 Tf 1 Tc 20 100 Td (AB) Tj ET";
        if (actualText) content += " EMC";
        PdfReadDocument document = PdfReadDocument.Open(BuildSingleStreamPdf(content, "<< /Font << /F1 5 0 R >> >>", font),
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = actualText ? 2 : 1 } });
        // ActualText replaces extraction, while the display retains the two original glyphs.
        // Zero-width glyphs use the aggregate fallback and consume no expansion budget.
        Assert.Equal(actualText ? 2 : 1, document.Pages[0].ToDrawing().Elements.OfType<OfficeDrawingText>().Count());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RenderPage_SpacedGlyphBudgetIsSharedAcrossRunsAndForms(bool splitRuns, bool useForm) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        string content = "BT /F1 10 Tf 1 Tc 20 100 Td " +
            (splitRuns ? "(AB) Tj (CD) Tj" : "(ABCD) Tj") + " ET";
        byte[] pdf = useForm
            ? BuildSingleStreamPdf("/Fm Do q 1 0 0 1 1000 0 cm /Fm Do Q /Fm Do", "<< /XObject << /Fm 6 0 R >> >>", font,
                BuildStreamObject(6, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources " + resources, content))
            : BuildSingleStreamPdf(content, resources, font);
        int exactCost = useForm ? 8 : 4;
        PdfReadDocument limited = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = exactCost - 1 }
        });
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => limited.Pages[0].ToDrawing());
        Assert.Equal(PdfReadLimitKind.PositionedTextCharacters, exception.Kind);
        Assert.Equal(exactCost - 1, exception.Limit);
        Assert.Equal(exactCost, exception.Actual);

        PdfReadDocument allowed = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = exactCost }
        });
        Assert.NotNull(allowed.Pages[0].ToDrawing());
        Assert.NotNull(allowed.Pages[0].ToDrawing()); // The budget belongs to one render invocation.
    }

    [Fact]
    public void RenderPage_DefaultBudgetRejectsOversizedSpacedRun() {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        string text = new string('A', PdfReadLimits.Default.MaxPositionedTextCharactersPerPage + 1);
        byte[] pdf = BuildSingleStreamPdf("BT /F1 10 Tf -6 Tc 20 100 Td (" + text + ") Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfPageImageRenderer.RenderPage(pdf));
        Assert.Equal(PdfReadLimitKind.PositionedTextCharacters, exception.Kind);
        Assert.Equal(text.Length, exception.Actual);
    }

    [Theory]
    [InlineData("87,0", -0.5, 0, false, false)]
    [InlineData("73%", 2, 0, false, false)]
    [InlineData("A B", 0, 4, false, false)]
    [InlineData("87,0", -0.5, 0, true, false)]
    [InlineData("A B", 2, 4, true, false)]
    [InlineData("87,0", -0.5, 0, false, true)]
    [InlineData("A B", 2, 4, true, true)]
    [InlineData("87,0", -8, 0, false, false)]
    [InlineData("87,0", -6, 0, false, false)]
    public void RenderPage_SpacedTextMatchesIndividuallyPositionedGlyphs(
        string value, double characterSpacing, double wordSpacing, bool clipped, bool rotated) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        string clip = clipped ? "90 80 26 30 re W n " : string.Empty;
        string matrix = rotated ? "0 1 -1 0 100 100 Tm " : "1 0 0 1 100 100 Tm ";
        string stream = clip + "BT /F1 10 Tf " + matrix +
            FormattableString.Invariant($"{characterSpacing} Tc {wordSpacing} Tw ({value}) Tj ET");

        // Courier has a 600-unit advance. Spacing moves the next glyph, while each
        // painted glyph remains six points wide. Explicit origins form the reference.
        var reference = new StringBuilder(clip);
        double offset = 0;
        foreach (char character in value) {
            reference.Append("BT /F1 10 Tf ").Append(matrix)
                .Append(offset.ToString("R", CultureInfo.InvariantCulture)).Append(" 0 Td (")
                .Append(character).Append(") Tj ET ");
            offset += 6 + characterSpacing + (character == ' ' ? wordSpacing : 0);
        }

        OfficeDrawing actual = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(stream, resources, font));
        OfficeDrawing expected = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(reference.ToString(), resources, font));
        byte[] expectedPixels = OfficeDrawingRasterRenderer.Render(expected).GetPixels();
        Assert.Contains(expectedPixels, channel => channel != 0);
        Assert.Equal(expectedPixels, OfficeDrawingRasterRenderer.Render(actual).GetPixels());
    }

    [Fact]
    public void ExportImage_ExcelSummaryKeepsEveryDigitAndPercentage() {
        // Desktop Excel PDF from the documented two-service operational dashboard example.
        string path = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Rendering", "excel-operational-summary.pdf");
        PdfReadDocument document = PdfReadDocument.Open(File.ReadAllBytes(path));
        Assert.Contains("87,0", document.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("73%", document.Pages[0].ExtractText(), StringComparison.Ordinal);

        var exported = document.Pages[0].ExportImage(OfficeImageExportFormat.Png);
        Assert.True(OfficePngReader.TryDecode(exported.Bytes, out OfficeRasterImage? image));
        foreach (var bounds in new[] {
            (Left: 221, Top: 86, Right: 225, Bottom: 95),
            (Left: 225, Top: 86, Right: 229, Bottom: 95),
            (Left: 229, Top: 86, Right: 231, Bottom: 95),
            (Left: 231, Top: 86, Right: 235, Bottom: 95),
            (Left: 221, Top: 107, Right: 225, Bottom: 115),
            (Left: 225, Top: 107, Right: 229, Bottom: 115),
            (Left: 229, Top: 107, Right: 236, Bottom: 115)
        }) {
            int painted = 0;
            for (int y = bounds.Top; y < bounds.Bottom; y++) {
                for (int x = bounds.Left; x < bounds.Right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A > 0 && pixel.R < 180 && pixel.G < 180 && pixel.B < 180) painted++;
                }
            }
            Assert.True(painted > 0, $"Missing glyph in {bounds}.");
        }
    }
}
