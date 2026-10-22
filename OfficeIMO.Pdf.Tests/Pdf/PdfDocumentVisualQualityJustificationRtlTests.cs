using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Letter = UglyToad.PdfPig.Content.Letter;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void CharacterJustificationDistributesDistinctRtlGlyphsWithoutChangingTheirVisualOrder() {
        const string text = "\u05D0\u05D1\u05D2\u05D3";
        PdfOptions options = JustificationPolicyOptions().UseFontFamily(new PdfEmbeddedFontFamily("Rtl Distribution",
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x05D0, 0x05D1, 0x05D2, 0x05D3)));
        options.TextShapingMode = PdfTextShapingMode.UnicodeScalar;
        var runs = new[] { new PdfTextRun(text).WithTextDirection(OfficeTextDirection.RightToLeft) };
        Letter[] control = ReadJustificationPolicyLetters(RenderJustificationPolicyRuns(options, runs, distributed: false))
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Letter[] actual = ReadJustificationPolicyLetters(RenderJustificationPolicyRuns(options, runs, distributed: true))
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Assert.Equal(4, control.Length);
        Assert.Equal(control.Select(letter => letter.Value), actual.Select(letter => letter.Value));
        double natural = control.Last().EndBaseLine.X - control.First().StartBaseLine.X;
        double extra = (240D - natural) / (control.Length - 1);
        for (int index = 1; index < actual.Length; index++)
            Assert.InRange(Math.Abs(actual[index].StartBaseLine.X - actual[index - 1].StartBaseLine.X -
                (control[index].StartBaseLine.X - control[index - 1].StartBaseLine.X) - extra), 0D, .02D);
        AssertJustificationPolicyWidth(actual.ToList(), expanded: true);
    }
}
