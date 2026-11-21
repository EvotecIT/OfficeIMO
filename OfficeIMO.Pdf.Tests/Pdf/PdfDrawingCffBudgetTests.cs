using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingCffBudgetTests {
    private const string Family = "Drawing CFF budget fixture";
    private const string Text = "The quick brown fox jumps over the lazy dog 0123456789 Sphinx of black quartz, judge my vow. Pack my box with five dozen liquor jugs!";

    [Fact]
    public void Drawings_sharing_one_cff_face_each_measure_within_their_own_charstring_budget() {
        string? path = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(path);
        const int drawings = 800;

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = OfficeLength.Points(800), PageHeight = OfficeLength.Points(400),
            MarginLeft = OfficeLength.Zero, MarginRight = OfficeLength.Zero,
            MarginTop = OfficeLength.Zero, MarginBottom = OfficeLength.Zero
        }.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(Family, File.ReadAllBytes(path!))))
            .Canvas(canvas => {
                // Every drawing measures the same face with distinct text, so the document executes
                // far more CharString operations than one drawing's allowance.
                for (int i = 0; i < drawings; i++) {
                    var drawing = new OfficeDrawing(800, 20).AddRichText(
                        new[] { new OfficeRichTextRun(Text + " #" + i.ToString(System.Globalization.CultureInfo.InvariantCulture), 9, OfficeColor.Black, fontFamily: Family) }, 0, 0, 800, 20,
                        wrapText: false);
                    // Publisher frames bound their clip by the measured glyph outlines.
                    ((OfficeDrawingRichText)drawing.Elements[0]).NativePublisherFrame = true;
                    canvas.Drawing(drawing, OfficeLength.Zero, OfficeLength.Points(i % 20 * 20),
                        OfficeLength.Points(800), OfficeLength.Points(20));
                }
            }).ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => string.Concat(page.Letters.Select(letter => letter.Value))));
        Assert.Contains("#0", text, StringComparison.Ordinal);
        Assert.Contains("#" + (drawings - 1).ToString(System.Globalization.CultureInfo.InvariantCulture), text, StringComparison.Ordinal);
    }
}
