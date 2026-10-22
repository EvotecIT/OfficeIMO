using OfficeIMO.Pdf;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(PdfTabAlignment.Right, PdfJustificationMode.Words)]
    [InlineData(PdfTabAlignment.Center, PdfJustificationMode.Words)]
    [InlineData(PdfTabAlignment.DecimalSeparator, PdfJustificationMode.Words)]
    [InlineData(PdfTabAlignment.Right, PdfJustificationMode.Characters)]
    [InlineData(PdfTabAlignment.Center, PdfJustificationMode.Characters)]
    [InlineData(PdfTabAlignment.DecimalSeparator, PdfJustificationMode.Characters)]
    public void JustificationPreservesNonLeftTabAnchorsOnExplicitBreakAndFinalLines(
        PdfTabAlignment alignment, PdfJustificationMode mode) {
        byte[] Render(PdfAlign paragraphAlignment, bool justifyBreaks, bool justifyFinal) {
            var style = new PdfParagraphStyle {
                JustificationMode = mode,
                JustifyExplicitLineBreaks = justifyBreaks,
                JustifyLastLine = justifyFinal
            };
            style.AddTabStop(160D, alignment);
            return PdfDocument.Create(JustificationPolicyOptions()).Paragraph(paragraph => paragraph
                .Text("Name\tUSD 12.50 net").LineBreak().Text("Name\tUSD 12.50 net"),
                paragraphAlignment, style: style).ToBytes();
        }

        using var natural = PdfPigDocument.Open(Render(PdfAlign.Left, true, true));
        using var policyOff = PdfPigDocument.Open(Render(PdfAlign.Justify, false, false));
        using var justified = PdfPigDocument.Open(Render(PdfAlign.Justify, true, true));
        List<List<Letter>> expected = GetNonWhitespaceLetterLines(natural.GetPage(1));
        List<List<Letter>> disabled = GetNonWhitespaceLetterLines(policyOff.GetPage(1));
        List<List<Letter>> actual = GetNonWhitespaceLetterLines(justified.GetPage(1));
        Assert.Equal(2, expected.Count);
        Assert.Equal(expected.Count, disabled.Count);
        Assert.Equal(expected.Count, actual.Count);
        for (int lineIndex = 0; lineIndex < actual.Count; lineIndex++) {
            List<Letter> controlLine = expected[lineIndex];
            List<Letter> disabledLine = disabled[lineIndex];
            List<Letter> actualLine = actual[lineIndex];
            Assert.Equal("NameUSD12.50net", string.Concat(actualLine.Select(letter => letter.Value)));
            Assert.Equal(controlLine.Count, disabledLine.Count);
            Assert.Equal(controlLine.Count, actualLine.Count);
            for (int letterIndex = 0; letterIndex < controlLine.Count; letterIndex++) {
                Letter control = controlLine[letterIndex];
                foreach (Letter candidate in new[] { disabledLine[letterIndex], actualLine[letterIndex] }) {
                    Assert.Equal(control.Value, candidate.Value);
                    Assert.Equal(control.StartBaseLine.X, candidate.StartBaseLine.X, 3);
                    Assert.Equal(control.EndBaseLine.X, candidate.EndBaseLine.X, 3);
                    Assert.Equal(control.StartBaseLine.Y, candidate.StartBaseLine.Y, 3);
                }
            }
            // Independently check the actual authored anchor, rather than only
            // comparing two artifacts that could share the same tab-placement bug.
            Letter[] field = actualLine.Skip("Name".Length).ToArray();
            double anchor = alignment switch {
                PdfTabAlignment.Center => (field.First().StartBaseLine.X + field.Last().EndBaseLine.X) / 2D,
                PdfTabAlignment.DecimalSeparator => Assert.Single(field, letter => letter.Value == ".").StartBaseLine.X,
                _ => field.Last().EndBaseLine.X
            };
            Assert.Equal(190D, anchor, 2);
        }
    }
}
