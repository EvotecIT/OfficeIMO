using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    private static (int ReferencePage, Letter Reference, int NotePage, Letter NoteLabel) AssertPdfFootnoteAssociation(
        PdfPigDocument pdf, string referenceText, string noteText, string label) {
        Assert.Single(label);
        Page referencePage = Assert.Single(pdf.GetPages(), page => page.Text.Contains(referenceText + label, StringComparison.Ordinal));
        Letter[] reference = AnchoredRunningLetters(referencePage, referenceText + label);
        Letter referenceMarker = reference[referenceText.Length];
        Letter referenceTail = reference[referenceText.Length - 1];
        Assert.Equal(label, referenceMarker.Value);
        Assert.True(referenceMarker.PointSize < referenceTail.PointSize);
        Assert.InRange(referenceMarker.StartBaseLine.Y - referenceTail.StartBaseLine.Y, double.Epsilon, referenceTail.PointSize);
        Assert.InRange(Math.Abs(referenceMarker.StartBaseLine.X - referenceTail.EndBaseLine.X), 0D, .15D);

        Page notePage = Assert.Single(pdf.GetPages(), page => page.Text.Contains(noteText, StringComparison.Ordinal));
        Letter[] text = AnchoredRunningLetters(notePage, noteText);
        Letter noteMarker = Assert.Single(notePage.Letters, letter => letter.Value == label &&
            letter.StartBaseLine.Y > text[0].StartBaseLine.Y &&
            letter.StartBaseLine.Y - text[0].StartBaseLine.Y < text[0].PointSize &&
            Math.Abs(letter.EndBaseLine.X - text[0].StartBaseLine.X) <= .15D);
        Assert.True(noteMarker.PointSize < text[0].PointSize);
        Assert.Equal(referencePage.Number, notePage.Number);
        Assert.True(referenceMarker.StartBaseLine.Y > noteMarker.StartBaseLine.Y);
        // AddFootNote authors a separate raised marker followed directly by
        // the supplied text. Association follows those glyphs and their source
        // positions, without requiring a space from the old flattened output.
        return (referencePage.Number, referenceMarker, notePage.Number, noteMarker);
    }

    private static void AssertPdfFootnoteSourceOrder(
        (int ReferencePage, Letter Reference, int NotePage, Letter NoteLabel) first,
        (int ReferencePage, Letter Reference, int NotePage, Letter NoteLabel) second) {
        Assert.True(PdfFootnoteGlyphPrecedes(first.ReferencePage, first.Reference, second.ReferencePage, second.Reference),
            "Footnote references must retain their source order.");
        Assert.True(PdfFootnoteGlyphPrecedes(first.NotePage, first.NoteLabel, second.NotePage, second.NoteLabel),
            "Footnote bodies must retain their reference order.");
    }

    private static bool PdfFootnoteGlyphPrecedes(int firstPage, Letter first, int secondPage, Letter second) =>
        firstPage < secondPage || firstPage == secondPage &&
        (first.StartBaseLine.Y > second.StartBaseLine.Y + .15D ||
            Math.Abs(first.StartBaseLine.Y - second.StartBaseLine.Y) <= .15D && first.StartBaseLine.X < second.StartBaseLine.X);
}
