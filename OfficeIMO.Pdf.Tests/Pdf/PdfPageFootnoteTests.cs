using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    private static PdfCreateOptions Options() => new() {
        BackgroundColor = null,
        PageWidth = 260, PageHeight = 140, MarginLeft = 20, MarginRight = 20,
        MarginTop = 20, MarginBottom = 20, DefaultFontSize = 8, DefaultFont = PdfStandardFont.Helvetica
    };

    private static PdfParagraphStyle Style() => new() {
        FontSize = 8, LineSpacing = PdfLineSpacing.Exactly(10), SpacingBefore = 0, SpacingAfter = 0
    };

    [Fact]
    public void ReferenceKeepsFirstFragmentAndGreedilyReservesNoteBeforeFollowingBody() {
        var doc = PdfDocument.Create(Options());
        var content = doc.Content;
        var note = content.DefineFootnote(body => {
            for (int index = 0; index < 3; index++) body.Text("NoteFragment" + index, style: Style());
        });
        content.Spacer(70).Paragraph(p => p.Text("Reference").Runs(new[] { PdfTextRun.Superscript("1").WithFootnoteReference(note) }), style: Style());
        content.Text("FollowingBody", style: Style());
        using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Reference", pdf.GetPage(1).Text);
        Assert.Contains("NoteFragment0", pdf.GetPage(1).Text);
        Assert.Contains("NoteFragment1", pdf.GetPage(1).Text);
        Assert.DoesNotContain("NoteFragment2", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingBody", pdf.GetPage(1).Text);
        Assert.Contains("NoteFragment2", pdf.GetPage(2).Text);
        Assert.Contains("FollowingBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void LongRichNoteUsesContinuationPagesAndBodyResumesOnFinalNotePage() {
        var doc = PdfDocument.Create(Options());
        var note = doc.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++)
                body.Paragraph(p => p.Bold("LongNote" + index.ToString("00")), style: Style());
        });
        doc.Content.Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        doc.Content.Text("FollowingBody", style: Style());
        using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("LongNote08", pdf.GetPage(1).Text);
        Assert.DoesNotContain("LongNote09", pdf.GetPage(1).Text);
        Assert.Contains("LongNote09", pdf.GetPage(2).Text);
        Assert.Contains("LongNote18", pdf.GetPage(2).Text);
        Assert.DoesNotContain("FollowingBody", pdf.GetPage(2).Text);
        Assert.Contains("LongNote24", pdf.GetPage(3).Text);
        Assert.Contains("FollowingBody", pdf.GetPage(3).Text);
        Assert.All(pdf.GetPages().SelectMany(page => page.Letters).Where(letter => !string.IsNullOrWhiteSpace(letter.Value)),
            letter => Assert.True(letter.StartBaseLine.Y >= 19D));
    }

    [Fact]
    public void ReferenceMovesWithItsNoteWhenFirstFragmentDoesNotFit() {
        var doc = PdfDocument.Create(Options());
        var note = doc.Content.DefineFootnote(body => body.Text("FootnoteText", style: Style()));
        doc.Content.Text("PriorBody", style: Style()).Spacer(75);
        doc.Content.Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("Reference", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FootnoteText", pdf.GetPage(1).Text);
        Assert.Contains("Reference", pdf.GetPage(2).Text);
        Assert.Contains("FootnoteText", pdf.GetPage(2).Text);
    }

    [Fact]
    public void EmptyReferencePreservesAssociationWithoutVisibleReplacementAndSerializesRepeatedly() {
        var doc = PdfDocument.Create(Options());
        var note = doc.Content.DefineFootnote(body => body.Text("AssociatedNote", style: Style()));
        doc.Content.Paragraph(p => p.Text("BodyText").Runs(new[] { PdfTextRun.Normal("").WithFootnoteReference(note).WithCharacterSpacing(1).WithTextCase(OfficeIMO.Drawing.OfficeTextCase.Uppercase) }), style: Style());
        foreach (byte[] bytes in new[] { doc.ToBytes(), doc.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            Assert.Single(pdf.GetPages());
            Assert.Equal("BodyTextAssociatedNote", pdf.GetPage(1).Text);
        }
    }

    [Fact]
    public void DefinitionIsDocumentBoundAndRejectsRecursiveNotes() {
        var owner = PdfDocument.Create(Options());
        var note = owner.Content.DefineFootnote(body => body.Text("Note", style: Style()));
        var other = PdfDocument.Create(Options());
        Assert.Throws<ArgumentException>(() => other.Content.Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("1").WithFootnoteReference(note) })));
        Assert.Throws<NotSupportedException>(() => owner.Content.DefineFootnote(body => body.Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("2").WithFootnoteReference(note) }))));
    }

    [Fact]
    public void ImpossibleFirstFragmentAndContinuationPageLimitFailWithoutUnboundedPagination() {
        var doc = PdfDocument.Create(Options());
        var note = doc.Content.DefineFootnote(body => body.Text("OversizedNote", style: new PdfParagraphStyle {
            FontSize = 8, LineSpacing = PdfLineSpacing.Exactly(110), SpacingAfter = 0
        }));
        doc.Content.Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        Assert.Throws<ArgumentException>(() => doc.ToBytes());

        var limitedOptions = Options();
        limitedOptions.MaxGeneratedPages = 2;
        var limited = PdfDocument.Create(limitedOptions);
        var longNote = limited.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++) body.Text("LongNote" + index, style: Style());
        });
        limited.Content.Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(longNote) }), style: Style());
        Assert.Throws<InvalidDataException>(() => limited.ToBytes());
    }
}
