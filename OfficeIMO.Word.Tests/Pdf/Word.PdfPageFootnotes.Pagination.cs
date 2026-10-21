using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    [Theory]
    [InlineData(24, 1, 1, 1, 1)]
    [InlineData(25, 1, 1, 1, 2)]
    [InlineData(27, 2, 2, 2, 2)]
    public void SaveAsPdf_PageFootnoteReservesSpaceBeforeCommittingReferenceAndBody(
        int fillerLines, int referencePage, int firstNotePage, int secondNotePage, int thirdNotePage) {
        var bytes = ConvertPageFootnoteFixture($"richnote-near-break-fill{fillerLines}");
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        int actualReferencePage = PageFootnoteTokenPage(actual, "NEAR REFERENCE");
        string[] paragraphs = Enumerable.Range(1, 3).Select(line => $"NOTE SHORT {line}").ToArray();
        Assert.Equal(actualReferencePage, PageFootnoteTokenPage(actual, paragraphs[0]));
        Assert.True(PageFootnoteTokenPage(actual, "FOLLOW NEAR") >= actualReferencePage);
        AssertPageFootnoteParagraphOrder(actual, paragraphs, bytes.Frame);
        foreach (string token in paragraphs)
            AssertPageFootnotePortableStyle(actual.GetPage(PageFootnoteTokenPage(actual, token)), token);
        string[] fillers = Enumerable.Range(1, fillerLines).Select(line => $"FILL{line:00}").ToArray();
        foreach (string token in fillers) AssertPageFootnoteTokenOnce(actual, token);
        AssertPageFootnoteBodyDoesNotOverlapNotes(actual, fillers.Concat(new[] { "NEAR REFERENCE", "FOLLOW NEAR" }), paragraphs);
        int[] pages = { firstNotePage, secondNotePage, thirdNotePage };
        AssertPageFootnoteNativeProfile(bytes, () => {
            Assert.Equal(2, reference.NumberOfPages);
            Assert.Equal(reference.NumberOfPages, actual.NumberOfPages);
            Assert.Equal(referencePage, actualReferencePage);
            Assert.Equal(2, PageFootnoteTokenPage(actual, "FOLLOW NEAR"));
            for (int line = 1; line <= 3; line++) {
                string token = paragraphs[line - 1];
                Assert.Equal(pages[line - 1], PageFootnoteTokenPage(reference, token));
                Assert.Equal(pages[line - 1], PageFootnoteTokenPage(actual, token));
                AssertPageFootnoteNativeText(actual.GetPage(pages[line - 1]), reference.GetPage(pages[line - 1]), token,
                    PageFootnoteNativeAnchor.PhysicalBottom);
            }
            var referenceGlyphs = PageFootnoteToken(actual.GetPage(referencePage), "NEAR REFERENCE");
            var firstNote = PageFootnoteToken(actual.GetPage(firstNotePage), "NOTE SHORT 1");
            Assert.True(referenceGlyphs[0].StartBaseLine.Y > firstNote[0].StartBaseLine.Y);
        });
    }

    [Fact]
    public void SaveAsPdf_PageFootnoteContinuesRichParagraphsOnceAndResumesBodyOnFinalPage() {
        var bytes = ConvertPageFootnoteFixture("richnote-long-rich-continuation");
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        Assert.Equal(1, PageFootnoteTokenPage(actual, "LONG REFERENCE"));
        string[] paragraphs = Enumerable.Range(1, 70).Select(number => $"NOTE LONG {number:000}").ToArray();
        Assert.Equal(1, PageFootnoteTokenPage(actual, paragraphs[0]));
        AssertPageFootnoteParagraphOrder(actual, paragraphs, bytes.Frame);
        for (int number = 1; number <= 70; number++) {
            string token = paragraphs[number - 1];
            AssertPageFootnotePortableStyle(actual.GetPage(PageFootnoteTokenPage(actual, token)), token,
                bold: number % 3 == 1, color: number % 3 == 0 ? 0x204080 : 0);
        }
        int finalNotePage = PageFootnoteTokenPage(actual, paragraphs[69]);
        string[] followingBody = Enumerable.Range(1, 6).Select(number => $"AFTER LONG {number:00}").ToArray();
        for (int number = 1; number <= 6; number++) {
            string token = followingBody[number - 1];
            AssertPageFootnoteTokenOnce(actual, token);
            Assert.Equal(finalNotePage, PageFootnoteTokenPage(actual, token));
        }
        AssertPageFootnoteBodyDoesNotOverlapNotes(actual, new[] { "BEGIN LONG", "LONG REFERENCE" }.Concat(followingBody), paragraphs);
        // Only the first-page separator is asserted. This fixture does not
        // prescribe universal selection of custom continuation separators.
        Assert.Contains("START SEPARATOR", actual.GetPage(1).Text, StringComparison.Ordinal);
        AssertPageFootnoteNativeProfile(bytes, () => {
            Assert.Equal(3, reference.NumberOfPages);
            Assert.Equal(reference.NumberOfPages, actual.NumberOfPages);
            for (int number = 1; number <= 70; number++) {
                string token = paragraphs[number - 1];
                int expectedPage = number <= 29 ? 1 : number <= 61 ? 2 : 3;
                Assert.Equal(expectedPage, PageFootnoteTokenPage(reference, token));
                Assert.Equal(expectedPage, PageFootnoteTokenPage(actual, token));
                AssertPageFootnoteNativeText(actual.GetPage(expectedPage), reference.GetPage(expectedPage), token,
                    PageFootnoteNativeAnchor.PhysicalBottom);
            }
            foreach (string token in followingBody) Assert.Equal(3, PageFootnoteTokenPage(actual, token));
            var finalPage = actual.GetPage(3);
            Assert.True(PageFootnoteToken(finalPage, "AFTER LONG 06")[0].StartBaseLine.Y >
                PageFootnoteToken(finalPage, "NOTE LONG 062")[0].StartBaseLine.Y);
        });
    }

    [Fact]
    public void SaveAsPdf_PageFootnoteBelongsToActualLateTableFragmentAndBottomFrame() {
        var bytes = ConvertPageFootnoteFixture("richnote-table-late-fragment");
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        int actualReferencePage = PageFootnoteTokenPage(actual, "TABLE REFERENCE");
        string[] paragraphs = { "NOTE TABLE FIRST", "NOTE TABLE LAST" };
        Assert.Equal(actualReferencePage, PageFootnoteTokenPage(actual, paragraphs[0]));
        foreach (string token in paragraphs) {
            AssertPageFootnoteTokenOnce(actual, token);
            Assert.Equal(actualReferencePage, PageFootnoteTokenPage(actual, token));
            AssertPageFootnotePortableStyle(actual.GetPage(actualReferencePage), token);
        }
        for (int line = 1; line <= 42; line++) {
            if (line != 34) AssertPageFootnoteTokenOnce(actual, $"CELL LINE {line:00}");
        }
        string[] cellTokens = Enumerable.Range(1, 42).Where(line => line != 34).Select(line => $"CELL LINE {line:00}").ToArray();
        AssertPageFootnoteParagraphOrder(actual, paragraphs, bytes.Frame);
        AssertPageFootnoteBodyDoesNotOverlapNotes(actual,
            new[] { "BEFORE TABLE", "TABLE REFERENCE", "AFTER TABLE" }.Concat(cellTokens), paragraphs);
        AssertPageFootnoteNativeProfile(bytes, () => {
            Assert.Equal(2, actual.NumberOfPages);
            Assert.Equal(2, actualReferencePage);
            foreach (string token in paragraphs)
                AssertPageFootnoteNativeText(actual.GetPage(2), reference.GetPage(2), token, PageFootnoteNativeAnchor.PhysicalBottom);
            var page = actual.GetPage(2);
            double noteStart = PageFootnoteToken(page, "NOTE TABLE FIRST")[0].StartBaseLine.Y;
            double noteEnd = PageFootnoteToken(page, "NOTE TABLE LAST")[0].StartBaseLine.Y;
            Assert.Equal(12D, noteStart - noteEnd, 3);
            Assert.True(PageFootnoteToken(page, "CELL LINE 42")[0].StartBaseLine.Y > noteStart);
            Assert.True(PageFootnoteToken(page, "AFTER TABLE")[0].StartBaseLine.Y > noteStart);
            Assert.True(noteEnd > PageFootnoteToken(page, "FOOTER")[0].StartBaseLine.Y);
        });
    }
}
