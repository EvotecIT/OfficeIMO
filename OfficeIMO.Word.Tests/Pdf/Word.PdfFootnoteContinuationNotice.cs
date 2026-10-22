using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    private const string FootnoteNoticeToken = "QUALIFIED CONTINUATION NOTICE";

    [Theory]
    [InlineData("14")]
    [InlineData("15")]
    public void SaveAsPdf_RegisteredFootnoteNoticeWithKeepNextAppearsOnceOnlyWhenNotesContinue(string mode) {
        using WordDocument document = LoadFootnoteStorySelectionFixture();
        SetFootnoteStorySelectionMode(document, mode);
        W.Footnote noticeStory = AddFootnoteNoticeStory(document, registered: true, empty: false);
        noticeStory.Elements<W.Paragraph>().Single().ParagraphProperties!.KeepNext = new W.KeepNext();

        using var pdf = PdfPigDocument.Open(ConvertFootnoteStorySelection(document));
        AssertFootnoteStorySelectionInventory(pdf);
        int finalNotePage = PageFootnoteTokenPage(pdf, "NOTE LONG 070");
        Assert.True(finalNotePage > 1);
        for (int number = 1; number <= pdf.NumberOfPages; number++) {
            var page = pdf.GetPage(number);
            if (number < finalNotePage) {
                var notice = PageFootnoteToken(page, FootnoteNoticeToken);
                Assert.All(notice, letter => Assert.Equal(12D, letter.PointSize, 3));
                string[] pageNotes = Enumerable.Range(1, 70).Select(index => $"NOTE LONG {index:000}")
                    .Where(token => PageFootnoteTokenPage(pdf, token) == number).ToArray();
                Assert.NotEmpty(pageNotes);
                Assert.True(PageFootnoteToken(page, pageNotes.Last())[0].StartBaseLine.Y > notice[0].StartBaseLine.Y);
            } else Assert.DoesNotContain(PageFootnoteCompact(FootnoteNoticeToken),
                PageFootnoteCompact(page.Text), StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public void SaveAsPdf_EmptyOrUnregisteredFootnoteNoticeDoesNotInventNoticeContent(bool registered, bool empty) {
        using WordDocument document = LoadFootnoteStorySelectionFixture();
        SetFootnoteStorySelectionMode(document, "14");
        W.Footnote notice = AddFootnoteNoticeStory(document, registered, empty);
        const string unusedFamily = "Unused Note Notice Face";
        if (!registered) {
            W.Paragraph paragraph = notice.Elements<W.Paragraph>().Single();
            paragraph.ParagraphProperties!.KeepNext = new W.KeepNext();
            paragraph.Elements<W.Run>().Single().RunProperties!.RunFonts = new W.RunFonts {
                Ascii = unusedFamily, HighAnsi = unusedFamily
            };
        }

        using var pdf = PdfPigDocument.Open(ConvertFootnoteStorySelection(document, registered ? null : unusedFamily));
        Assert.True(pdf.NumberOfPages > 1);
        AssertFootnoteStorySelectionInventory(pdf);
        Assert.DoesNotContain(PageFootnoteCompact(FootnoteNoticeToken),
            PageFootnoteCompact(string.Concat(pdf.GetPages().Select(page => page.Text))), StringComparison.Ordinal);
    }

    private static W.Footnote AddFootnoteNoticeStory(WordDocument document, bool registered, bool empty) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        var paragraph = new W.Paragraph(new W.ParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto
        }));
        if (!empty) paragraph.Append(new W.Run(new W.RunProperties(new W.FontSize { Val = "24" }),
            new W.Text(FootnoteNoticeToken)));
        var notice = new W.Footnote(paragraph) {
            Id = -13, Type = W.FootnoteEndnoteValues.ContinuationNotice
        };
        main.FootnotesPart!.Footnotes!.PrependChild(notice);
        if (registered) main.DocumentSettingsPart!.Settings!.GetFirstChild<W.FootnoteDocumentWideProperties>()!
            .Append(new W.FootnoteSpecialReference { Id = -13 });
        return notice;
    }
}
