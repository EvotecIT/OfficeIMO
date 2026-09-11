using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingTextFitTests {
    public static System.Collections.Generic.IEnumerable<object[]> StandardFontCases() {
        foreach (PdfStandardFont font in System.Enum.GetValues(typeof(PdfStandardFont)))
            foreach (OfficeTextVerticalAlignment alignment in new[] { OfficeTextVerticalAlignment.Top,
                OfficeTextVerticalAlignment.Center, OfficeTextVerticalAlignment.Bottom })
                yield return new object[] { font, alignment };
    }

    [Theory]
    [MemberData(nameof(StandardFontCases))]
    public void UnembeddedStandardFontFitContainsAccentsAndDescenders(PdfStandardFont font, OfficeTextVerticalAlignment alignment) {
        string name = font.ToString();
        string family = name.StartsWith("Times") ? "Times New Roman" : name.StartsWith("Courier") ? "Courier" : "Helvetica";
        const string value = "\u00C1gypsy";
        var drawing = new OfficeDrawing(100, 50);
        drawing.AddRichText(new[] { new OfficeRichTextRun(value, 20, OfficeColor.Black,
            bold: name.Contains("Bold"), italic: name.Contains("Italic") || name.Contains("Oblique"), fontFamily: family) },
            10, 10, 80, 15, lineHeight: 10, verticalAlignment: alignment, shrinkToFit: true);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            DefaultFont = font, PageWidth = 100, PageHeight = 50, MarginLeft = 0, MarginTop = 0, MarginRight = 0, MarginBottom = 0
        }).Compose(c => c.Page(p => p.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(value, string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => {
            Assert.InRange(letter.BoundingBox.Bottom, 24.99D, 40.01D);
            Assert.InRange(letter.BoundingBox.Top, 24.99D, 40.01D);
        });
    }

    [Theory]
    [InlineData("gypsy", OfficeTextVerticalAlignment.Top)]
    [InlineData("gypsy", OfficeTextVerticalAlignment.Bottom)]
    [InlineData("\u00C1gj", OfficeTextVerticalAlignment.Center)]
    [InlineData("\u00C1gj", OfficeTextVerticalAlignment.Bottom)]
    public void CondensedFitRetainsActualGlyphBounds(string value, OfficeTextVerticalAlignment alignment) {
        var drawing = new OfficeDrawing(100, 50);
        drawing.Fonts.Add("Proof Sans", File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()!));
        drawing.AddRichText(new[] { new OfficeRichTextRun(value, 20, OfficeColor.Black, fontFamily: "Proof Sans") },
            10, 10, 80, 15, lineHeight: 10, verticalAlignment: alignment, shrinkToFit: true);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            PageWidth = 100, PageHeight = 50, MarginLeft = 0, MarginTop = 0, MarginRight = 0, MarginBottom = 0
        }.UseRenderingProfile(new OfficeRenderingProfile("fit", drawing.Fonts))).Compose(c => c.Page(p => p.Content(content => content.Drawing(drawing)))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(value, string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => {
            Assert.InRange(letter.BoundingBox.Bottom, 24.99D, 40.01D);
            Assert.InRange(letter.BoundingBox.Top, 24.99D, 40.01D);
        });
    }
}
