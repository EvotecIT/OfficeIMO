using System;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData("", false, false, 220D)]
    [InlineData(" ", false, false, 220D)]
    [InlineData(" ", false, true, 220D)]
    [InlineData("  ", true, false, 220D)]
    [InlineData("  ", true, true, 220D)]
    [InlineData("  ", true, false, 24D)]
    public void InvisibleReferenceLabelsRetainOneNoteWithoutChangingAuthoredWhitespace(string label, bool preserve, bool terminal, double width) {
        PdfCreateOptions options = Options();
        options.PageWidth = width + options.MarginLeft + options.MarginRight;
        options.DefaultFont = PdfStandardFont.Courier;
        options.TextWhitespaceMode = preserve ? PdfTextWhitespaceMode.Preserve : PdfTextWhitespaceMode.Collapse;
        var control = PdfDocument.Create(options);
        AddWhitespaceReferenceBody(control.Content, PdfTextRun.Normal(label), terminal);
        using var controlPdf = PdfPigDocument.Open(control.ToBytes());
        var expected = controlPdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();

        var document = PdfDocument.Create(options);
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Text("AssociatedNote", style: Style()));
        AddWhitespaceReferenceBody(document.Content, PdfTextRun.Normal(label).WithFootnoteReference(note), terminal);
        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(1, pdf.NumberOfPages);
            var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
            Assert.Equal(expected.Length + "AssociatedNote".Length, letters.Length);
            Assert.Equal(terminal ? "LEFT" : "LEFTRIGHT", string.Concat(letters.Take(expected.Length).Select(letter => letter.Value)));
            Assert.Single(Regex.Matches(string.Concat(letters.Select(letter => letter.Value)), "AssociatedNote"));
            for (int index = 0; index < expected.Length; index++) {
                Assert.InRange(Math.Abs(expected[index].StartBaseLine.X - letters[index].StartBaseLine.X), 0D, .02D);
                Assert.InRange(Math.Abs(expected[index].StartBaseLine.Y - letters[index].StartBaseLine.Y), 0D, .02D);
            }
        }
    }

    [Fact]
    public void LaterReferenceToSameDefinitionCannotPostponeWrappedWhitespaceReference() {
        PdfCreateOptions options = Options();
        options.PageWidth = options.MarginLeft + 24D + options.MarginRight;
        options.DefaultFont = PdfStandardFont.Courier;
        options.TextWhitespaceMode = PdfTextWhitespaceMode.Preserve;
        var document = PdfDocument.Create(options);
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Text("NOTE", style: Style()));
        document.Content.Paragraph(paragraph => {
            // Both references belong to this one wrapped paragraph. The spaces after
            // LEFT lose their painted carrier when RIGHT wraps out of the full line.
            paragraph.Text("LEFT").Runs(new[] { PdfTextRun.Normal("  ").WithFootnoteReference(note) }).Text("RIGHT").LineBreak();
            for (int index = 0; index < 12; index++) paragraph.Text("PAD").LineBreak();
            paragraph.Runs(new[] { PdfTextRun.Normal("LATE").WithFootnoteReference(note) });
        }, style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("LEFT", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("RIGHT", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("LATE", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("LATE", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("NOTE", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("NOTE", pdf.GetPage(2).Text, StringComparison.Ordinal);
        string painted = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Single(Regex.Matches(painted, "NOTE"));
    }

    private static void AddWhitespaceReferenceBody(PdfContentBuilder body, PdfTextRun marker, bool terminal) {
        body.Paragraph(paragraph => {
            paragraph.Text("LEFT").Runs(new[] { marker });
            if (!terminal) paragraph.Text("RIGHT");
        }, style: Style());
    }
}
