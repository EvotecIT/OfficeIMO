using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfPigPage = UglyToad.PdfPig.Content.Page;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfContinuedFlowCapacityTests {
    [Fact]
    public void ContinuationSeparatorMarginPolicyIsOptInAndKeepsFirstFrameRemainingCapacity() {
        PdfDocument strict = StoryDocument(32);
        PdfDocument enabled = StoryDocument(32, allowBottomMargin: true);

        using var strictPdf = PdfPigDocument.Open(strict.ToBytes());
        Assert.Equal(new[] { 8, 10, 10, 4 }, strictPdf.GetPages().Select(page => StoryTokens(page).Length));
        AssertStoryOrder(strictPdf, 32);
        foreach (byte[] bytes in new[] { enabled.ToBytes(), enabled.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(new[] { 8, 12, 12 }, pdf.GetPages().Select(page => StoryTokens(page).Length));
            Assert.Equal(StoryTokens(strictPdf.GetPage(1)), StoryTokens(pdf.GetPage(1)));
            Assert.Contains("BodyLead", pdf.GetPage(1).Text, StringComparison.Ordinal);
            AssertStoryOrder(pdf, 32);
            foreach (PdfPigPage page in pdf.GetPages().Skip(1)) {
                var separator = Assert.Single(page.GetWords(), word => word.Text == "ContinuationSeparator");
                var first = Assert.Single(page.GetWords(), word => word.Text == StoryTokens(page)[0]);
                Assert.True(separator.BoundingBox.Bottom > first.BoundingBox.Top);
            }
        }
    }

    [Fact]
    public void ContinuationSeparatorMarginCreditStopsAtThePhysicalPageEdge() {
        // A 40-point separator can occupy only the 20-point bottom margin.
        PdfDocument document = StoryDocument(28, allowBottomMargin: true, continuationHeight: 40);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(new[] { 8, 10, 10 }, pdf.GetPages().Select(page => StoryTokens(page).Length));
        AssertStoryOrder(pdf, 28);
        Assert.All(pdf.GetPages().SelectMany(page => page.GetWords()), word => Assert.True(word.BoundingBox.Bottom >= 0D));
    }

    [Fact]
    public void MeasuredBlankContinuationNoticeRetainsStrictSeparatorAndNoticeCapacity() {
        PdfDocument strict = StoryDocument(32, emptyNotice: true);
        PdfDocument enabled = StoryDocument(32, allowBottomMargin: true, emptyNotice: true);
        using var strictPdf = PdfPigDocument.Open(strict.ToBytes());
        using var pdf = PdfPigDocument.Open(enabled.ToBytes());
        Assert.Equal(new[] { 7, 9, 9, 7 }, pdf.GetPages().Select(page => StoryTokens(page).Length));
        Assert.Equal(strictPdf.GetPages().SelectMany(StoryTokens), pdf.GetPages().SelectMany(StoryTokens));
        Assert.Equal(strictPdf.GetPages().Select(page => StoryTokens(page).Length),
            pdf.GetPages().Select(page => StoryTokens(page).Length));
        AssertStoryOrder(pdf, 32);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectedSourceFooterProtectsOnlyItsActualVariantIncludingTextFallback(bool useTextFallback) {
        PdfDocument document = PdfDocument.Create(Options());
        document.Footer(footer => {
            footer.Text("").EvenPagesText("FooterFallback").FontSize(6).Offset(10);
            if (useTextFallback) footer.EvenPagesContent(content => content.Spacer(200), 10);
            else footer.EvenPagesContent(content => content.Text("FooterFlow", style: Style()), 10);
        });
        if (useTextFallback) document.Options.AllowRunningContentTextFallback(isHeader: false);
        AddStory(document, 30, allowBottomMargin: true);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(new[] { 8, 10, 12 }, pdf.GetPages().Select(page => StoryTokens(page).Length));
        AssertStoryOrder(pdf, 30);
        string footerToken = useTextFallback ? "FooterFallback" : "FooterFlow";
        var footerWord = Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == footerToken);
        Assert.DoesNotContain(footerToken, pdf.GetPage(3).Text, StringComparison.Ordinal);
        Assert.All(pdf.GetPage(2).GetWords().Where(word => word.Text.StartsWith("Story", StringComparison.Ordinal)),
            word => Assert.True(word.BoundingBox.Bottom > footerWord.BoundingBox.Top));
    }

    [Fact]
    public void AcceptedPageFootnotesKeepTheirReservationAboveTheContinuationSeparatorMargin() {
        PdfCreateOptions options = Options();
        options.PageHeight = 140;
        PdfDocument document = PdfDocument.Create(options);
        var tableStyle = new PdfTableStyle {
            HeaderRowCount = 0, RepeatHeaderRowCount = 0, MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0,
            FixedRowHeights = new List<double?> { 80D, 80D, 80D }, AllowRowBreakAcrossPages = false,
            FontSize = 6, CellPaddingX = 2, CellPaddingY = 0, SpacingBefore = 0, SpacingAfter = 0
        };
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => body.Table(new[] {
            new[] { "PageNote0" }, new[] { "PageNote1" }, new[] { "PageNote2" }
        }, style: tableStyle));
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: Style());
        document.Content.ContinuedFlow(content => WriteStory(content, 4), separator: null,
            continuationSeparator: content => content.Text("ContinuationSeparator", style: Style()),
            continuationNotice: null, reserveFinalNotice: false, allowContinuationSeparatorInBottomMargin: true);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(new[] { 1, 1, 1, 1 }, pdf.GetPages().Select(page => StoryTokens(page).Length));
        AssertStoryOrder(pdf, 4);
        for (int index = 0; index < 3; index++) {
            PdfPigPage page = pdf.GetPage(index + 1);
            var noteWord = Assert.Single(page.GetWords(), word => word.Text == "PageNote" + index);
            var storyWord = Assert.Single(page.GetWords(), word => word.Text == "Story" + index.ToString("00"));
            Assert.True(storyWord.BoundingBox.Bottom > noteWord.BoundingBox.Top);
            Assert.Single(pdf.GetPages().SelectMany(candidate => candidate.GetWords()), word => word.Text == noteWord.Text);
        }
    }

    private static PdfDocument StoryDocument(int count, bool allowBottomMargin = false,
        double continuationHeight = 20D, bool emptyNotice = false) {
        PdfDocument document = PdfDocument.Create(Options());
        AddStory(document, count, allowBottomMargin, continuationHeight, emptyNotice);
        return document;
    }

    private static void AddStory(PdfDocument document, int count, bool allowBottomMargin,
        double continuationHeight = 20D, bool emptyNotice = false) {
        document.Content.Text("BodyLead", style: Style(20D));
        Action<PdfContentBuilder> build = content => WriteStory(content, count);
        Action<PdfContentBuilder> first = content => content.Text("FirstSeparator", style: Style(20D));
        Action<PdfContentBuilder> later = content => content.Text("ContinuationSeparator", style: Style(continuationHeight));
        Action<PdfContentBuilder>? notice = emptyNotice ? content => content.Spacer(10D) : null;
        if (allowBottomMargin)
            document.Content.ContinuedFlow(build, first, later, notice, reserveFinalNotice: false,
                allowContinuationSeparatorInBottomMargin: true);
        else document.Content.ContinuedFlow(build, first, later, notice, reserveFinalNotice: false);
    }

    private static void WriteStory(PdfContentBuilder content, int count) {
        for (int index = 0; index < count; index++) content.Text("Story" + index.ToString("00"), style: Style());
    }

    private static string[] StoryTokens(PdfPigPage page) => page.GetWords()
        .Where(word => word.Text.StartsWith("Story", StringComparison.Ordinal)).Select(word => word.Text).ToArray();

    private static void AssertStoryOrder(PdfPigDocument pdf, int count) =>
        Assert.Equal(Enumerable.Range(0, count).Select(index => "Story" + index.ToString("00")), pdf.GetPages().SelectMany(StoryTokens));

    private static PdfCreateOptions Options() => new() {
        PageWidth = 260, PageHeight = 160, MarginTop = 20, MarginBottom = 20, MarginLeft = 20, MarginRight = 20,
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 6
    };

    private static PdfParagraphStyle Style(double height = 10D) => new() {
        FontSize = 6, LineSpacing = PdfLineSpacing.Exactly(height), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
    };
}
