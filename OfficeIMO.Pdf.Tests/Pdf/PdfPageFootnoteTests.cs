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
        PageSetup = {
            Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(260), global::OfficeIMO.OfficeLength.Points(140)),
            Margins = { Left = global::OfficeIMO.OfficeLength.Points(20), Right = global::OfficeIMO.OfficeLength.Points(20), Top = global::OfficeIMO.OfficeLength.Points(20), Bottom = global::OfficeIMO.OfficeLength.Points(20) }
        }, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(8), DefaultFont = PdfStandardFont.Helvetica
    };

    private static PdfParagraphStyle Style() => new() {
        FontSize = global::OfficeIMO.OfficeLength.Points(8), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(10)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
    };

    [Fact]
    public void ReferenceKeepsFirstFragmentAndGreedilyReservesNoteBeforeFollowingBody() {
        var doc = PdfDocument.Create(Options());
        var content = doc.Content;
        var note = content.DefineFootnote(body => {
            for (int index = 0; index < 3; index++) body.AddParagraph("NoteFragment" + index, style: Style());
        });
        content.Spacer(global::OfficeIMO.OfficeLength.Points(70)).AddParagraph(p => { p.AddText("Reference"); p.AddRuns(new[] { PdfTextRun.Superscript("1").WithFootnoteReference(note) }); }, style: Style());
        content.AddParagraph("FollowingBody", style: Style());
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
                body.AddParagraph(p => p.AddText("LongNote" + index.ToString("00")).SetBold(), style: Style());
        });
        doc.Content.AddParagraph(p => p.AddRuns(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        doc.Content.AddParagraph("FollowingBody", style: Style());
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
        var note = doc.Content.DefineFootnote(body => body.AddParagraph("FootnoteText", style: Style()));
        doc.Content.AddParagraph("PriorBody", style: Style()).Spacer(global::OfficeIMO.OfficeLength.Points(75));
        doc.Content.AddParagraph(p => p.AddRuns(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
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
        var note = doc.Content.DefineFootnote(body => body.AddParagraph("AssociatedNote", style: Style()));
        doc.Content.AddParagraph(p => { p.AddText("BodyText"); p.AddRuns(new[] { PdfTextRun.Normal("").WithFootnoteReference(note).WithCharacterSpacing(global::OfficeIMO.OfficeLength.Points(1)).WithTextCase(OfficeIMO.Drawing.OfficeTextCase.Uppercase) }); }, style: Style());
        foreach (byte[] bytes in new[] { doc.ToBytes(), doc.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            Assert.Single(pdf.GetPages());
            Assert.Equal("BodyTextAssociatedNote", pdf.GetPage(1).Text);
        }
    }

    [Fact]
    public void DefinitionIsDocumentBoundAndRejectsRecursiveNotes() {
        var owner = PdfDocument.Create(Options());
        var note = owner.Content.DefineFootnote(body => body.AddParagraph("Note", style: Style()));
        var other = PdfDocument.Create(Options());
        Assert.Throws<ArgumentException>(() => other.Content.AddParagraph(p => p.AddRuns(new[] { PdfTextRun.Normal("1").WithFootnoteReference(note) })));
        Assert.Throws<NotSupportedException>(() => owner.Content.DefineFootnote(body => body.AddParagraph(p => p.AddRuns(new[] { PdfTextRun.Normal("2").WithFootnoteReference(note) }))));
    }

    [Fact]
    public void ImpossibleFirstFragmentAndContinuationPageLimitFailWithoutUnboundedPagination() {
        var doc = PdfDocument.Create(Options());
        var note = doc.Content.DefineFootnote(body => body.AddParagraph("OversizedNote", style: new PdfParagraphStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(8), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(110)), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
        }));
        doc.Content.AddParagraph(p => p.AddRuns(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        Assert.Throws<ArgumentException>(() => doc.ToBytes());

        var limitedOptions = Options();
        limitedOptions.MaxGeneratedPages = 2;
        var limited = PdfDocument.Create(limitedOptions);
        var longNote = limited.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++) body.AddParagraph("LongNote" + index, style: Style());
        });
        limited.Content.AddParagraph(p => p.AddRuns(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(longNote) }), style: Style());
        Assert.Throws<InvalidDataException>(() => limited.ToBytes());
    }
}
