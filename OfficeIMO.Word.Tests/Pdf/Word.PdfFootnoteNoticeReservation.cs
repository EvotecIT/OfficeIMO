using System.Text.RegularExpressions;
using OfficeIMO.Word;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    [Theory]
    [InlineData("continuation-notice-compat14-font6", 14, 6, false, 3, 440.880, 6)]
    [InlineData("continuation-notice-compat14-font18", 14, 18, false, 4, 406.320, 2)]
    [InlineData("continuation-notice-compat14-short-note", 14, 12, true, 1, 433.920, 6)]
    [InlineData("continuation-notice-compat15-short-note", 15, 12, true, 1, 447.720, 6)]
    public void SaveAsPdf_FootnoteContinuationNoticeReservationUsesSourceMetricsAndCompatibility(
        string id, int mode, int noticeFontSize, bool shortNote, int expectedPages,
        double nativeFinalNoteBaselineTop, int followingBodyOnLastNotePage) {
        PageFootnoteArtifact bytes = ConvertFootnoteRuleFixture(id, typedOrdinaryNoteInventory: true);
        using var actual = PdfPigDocument.Open(bytes.Actual);
        using var reference = PdfPigDocument.Open(bytes.Reference);
        string[] noteTokens = shortNote ? new[] { "NOTE SHORT BODY" } :
            Enumerable.Range(1, 112).Select(number => $"NOTE {number:000}").ToArray();
        string[] following = Enumerable.Range(1, 6).Select(number => $"AFTER LONG {number:00}").ToArray();
        NoticeNoteInventory inventory = ReadNoticeNoteInventory(actual);
        NoticeNoteToken[] tokens = noteTokens.Select(inventory.Token).ToArray();
        int[] notePages = inventory.Glyphs.Select(glyph => glyph.Page).Distinct().ToArray();
        int lastNotePage = inventory.Glyphs.Last().Page;
        string notice = $"NOTICE MODE {mode} CONTINUED ON NEXT PAGE";
        AssertFootnoteRuleTokenOrder(actual, following);
        Assert.True(PageFootnoteTokenPage(actual, following[0]) >= lastNotePage);
        if (shortNote) AssertPageFootnotePortableStyle(actual.GetPage(notePages[0]), noteTokens[0]);
        foreach (Page page in actual.GetPages()) {
            bool hasNote = notePages.Contains(page.Number);
            bool continues = hasNote && page.Number < lastNotePage;
            Assert.Equal(continues, PageFootnoteCompact(page.Text).Contains(PageFootnoteCompact(notice), StringComparison.Ordinal));
            if (continues) {
                foreach (Letter letter in PageFootnoteToken(page, notice)) Assert.Equal((double)noticeFontSize, letter.PointSize, 3);
            }
            if (hasNote) {
                string expected = page.Number == notePages[0] || mode == 15 ? "START SEPARATOR" : "CONTINUATION SEPARATOR";
                string excluded = page.Number == notePages[0] || mode == 15 ? "CONTINUATION SEPARATOR" : "START SEPARATOR";
                PageFootnoteToken(page, expected);
                Assert.DoesNotContain(PageFootnoteCompact(excluded), PageFootnoteCompact(page.Text), StringComparison.Ordinal);
            } else {
                Assert.DoesNotContain("STARTSEPARATOR", PageFootnoteCompact(page.Text), StringComparison.Ordinal);
                Assert.DoesNotContain("CONTINUATIONSEPARATOR", PageFootnoteCompact(page.Text), StringComparison.Ordinal);
            }
        }
        AssertPageFootnoteNativeProfile(bytes, () => {
            Assert.Equal(expectedPages, reference.NumberOfPages);
            Assert.Equal(reference.NumberOfPages, actual.NumberOfPages);
            NoticeNoteInventory nativeInventory = ReadNoticeNoteInventory(reference);
            Assert.Equal(inventory.Text, nativeInventory.Text);
            NoticeNoteToken[] nativeTokens = noteTokens.Select(nativeInventory.Token).ToArray();
            // The independent pinned export owns every token's page assignment;
            // total page count alone would miss changes within the long note.
            // A word/number pair can span pages: NOTE085 does so in both pinned
            // font6 exports. Retain both ends instead of requiring a whole token
            // on one page or mixing notice/footer glyphs into its note stream.
            Assert.Equal(nativeTokens.Select(token => token.FirstPage), tokens.Select(token => token.FirstPage));
            Assert.Equal(nativeTokens.Select(token => token.EndPage), tokens.Select(token => token.EndPage));
            int[] nativeNotePages = nativeInventory.Glyphs.Select(glyph => glyph.Page).Distinct().ToArray();
            int nativeLastNotePage = nativeInventory.Glyphs.Last().Page;
            for (int index = 0; index < following.Length; index++) {
                string token = following[index];
                int nativeBodyPage = PageFootnoteTokenPage(reference, token);
                int expectedBodyPage = index < followingBodyOnLastNotePage ? nativeLastNotePage : nativeLastNotePage + 1;
                Assert.Equal(expectedBodyPage, nativeBodyPage);
                Assert.Equal(nativeBodyPage, PageFootnoteTokenPage(actual, token));
                AssertFootnoteRuleTextPosition(actual.GetPage(nativeBodyPage), reference.GetPage(nativeBodyPage), token);
            }
            foreach (Page nativePage in reference.GetPages()) {
                bool hasNote = nativeNotePages.Contains(nativePage.Number);
                bool continues = hasNote && nativePage.Number < nativeLastNotePage;
                Assert.Equal(continues, PageFootnoteCompact(nativePage.Text).Contains(PageFootnoteCompact(notice), StringComparison.Ordinal));
                if (continues) AssertFootnoteRuleTextPosition(actual.GetPage(nativePage.Number), nativePage, notice);
                if (!hasNote) continue;
                Page observedPage = actual.GetPage(nativePage.Number);
                Letter[] nativeLetters = nativeInventory.PageLetters(nativePage.Number);
                Letter[] observedLetters = inventory.PageLetters(nativePage.Number);
                Assert.NotEmpty(nativeLetters);
                Assert.NotEmpty(observedLetters);
                Assert.Equal(PageFootnoteCompact(string.Concat(nativeLetters.Select(letter => letter.Value))),
                    PageFootnoteCompact(string.Concat(observedLetters.Select(letter => letter.Value))));
                double nativeFirst = (double)nativePage.Height - nativeLetters.Max(letter => letter.StartBaseLine.Y);
                double nativeLast = (double)nativePage.Height - nativeLetters.Min(letter => letter.StartBaseLine.Y);
                double observedFirst = (double)observedPage.Height - observedLetters.Max(letter => letter.StartBaseLine.Y);
                double observedLast = (double)observedPage.Height - observedLetters.Min(letter => letter.StartBaseLine.Y);
                _output.WriteLine(FormattableString.Invariant(
                    $"{id}: page={nativePage.Number}; nativeNoteTop={nativeFirst:F6}..{nativeLast:F6}; actualNoteTop={observedFirst:F6}..{observedLast:F6}; noticeInk={continues}"));
                Assert.InRange(Math.Abs(observedFirst - nativeFirst), 0D, .20D);
                Assert.InRange(Math.Abs(observedLast - nativeLast), 0D, .20D);
                if (nativePage.Number == nativeLastNotePage)
                    Assert.InRange(Math.Abs(nativeLast - nativeFinalNoteBaselineTop), 0D, .001D);
            }
        });
    }

    private static NoticeNoteInventory ReadNoticeNoteInventory(PdfPigDocument pdf) {
        // AssertNoticeFixtureSourceProfile qualifies this interval before the
        // conversion. It selects the sole ordinary Arial10 story, not generic
        // small text: native Word serializes that source size as 9.96pt.
        var glyphs = new List<NoticeNoteGlyph>();
        foreach (Page page in pdf.GetPages()) {
            double priorY = double.PositiveInfinity;
            double priorX = double.NegativeInfinity;
            foreach (Letter letter in page.Letters.Where(letter => letter.PointSize >= 9.95D &&
                letter.PointSize <= 10.01D && !string.IsNullOrWhiteSpace(letter.Value))) {
                Assert.True(letter.StartBaseLine.Y < priorY - .01D ||
                    Math.Abs(letter.StartBaseLine.Y - priorY) <= .01D && letter.StartBaseLine.X >= priorX,
                    $"Ordinary note glyph '{letter.Value}' is out of source reading order on page {page.Number}.");
                priorY = letter.StartBaseLine.Y;
                priorX = letter.StartBaseLine.X;
                foreach (char character in letter.Value.Where(character => !char.IsWhiteSpace(character)))
                    glyphs.Add(new NoticeNoteGlyph(character, page.Number, letter));
            }
        }
        Assert.NotEmpty(glyphs);
        return new NoticeNoteInventory(glyphs);
    }

    private static void AssertNoticeFixtureSourceProfile(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Styles styles = main.StyleDefinitionsPart!.Styles!;
        W.Style normal = Assert.Single(styles.Elements<W.Style>(), style => style.StyleId?.Value == "Normal");
        W.Style note = Assert.Single(styles.Elements<W.Style>(), style => style.StyleId?.Value == "FootnoteText");
        W.Style label = Assert.Single(styles.Elements<W.Style>(), style => style.StyleId?.Value == "FootnoteReference");
        Assert.Equal("24", normal.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal("Arial", normal.StyleRunProperties.RunFonts!.Ascii!.Value);
        Assert.Equal("20", note.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal("Arial", note.StyleRunProperties.RunFonts!.Ascii!.Value);
        Assert.Equal("Arial", note.StyleRunProperties.RunFonts.HighAnsi!.Value);
        Assert.Equal(W.VerticalPositionValues.Superscript, label.StyleRunProperties!.VerticalTextAlignment!.Val!.Value);
        W.Footnote ordinary = Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(),
            item => item.Type == null || item.Type.Value == W.FootnoteEndnoteValues.Normal);
        W.Paragraph paragraph = Assert.Single(ordinary.Elements<W.Paragraph>());
        Assert.Equal("FootnoteText", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        foreach (W.Run run in paragraph.Elements<W.Run>()) {
            if (run.Elements<W.Text>().Any()) Assert.Null(run.RunProperties);
            if (run.Elements<W.FootnoteReferenceMark>().Any())
                Assert.Equal("FootnoteReference", run.RunProperties!.RunStyle!.Val!.Value);
        }
        // Body/footer visible runs inherit Normal12. Their only character-style
        // changes belong to superscript labels, so they cannot enter Arial10.
        foreach (var root in new DocumentFormat.OpenXml.OpenXmlElement[] { main.Document.Body! }
            .Concat(main.FooterParts.Select(part => (DocumentFormat.OpenXml.OpenXmlElement)part.Footer!))) {
            Assert.DoesNotContain(root.Descendants<W.ParagraphStyleId>(), style => style.Val?.Value != "Normal");
            foreach (W.Run run in root.Descendants<W.Run>().Where(run => run.Elements<W.Text>().Any()))
                Assert.Null(run.RunProperties);
        }
        foreach (W.Footnote special in main.FootnotesPart.Footnotes.Elements<W.Footnote>().Where(item => item != ordinary)) {
            foreach (W.Paragraph specialParagraph in special.Elements<W.Paragraph>()) {
                foreach (W.Run run in specialParagraph.Elements<W.Run>().Where(run => run.Elements<W.Text>().Any())) {
                    string? size = run.RunProperties?.FontSize?.Val?.Value ??
                        specialParagraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.FontSize>()?.Val?.Value ?? "24";
                    Assert.Contains(size, new[] { "12", "24", "36" }); // notices6/12/18; separator text12
                }
            }
        }
    }

    private sealed record NoticeNoteGlyph(char Value, int Page, Letter Letter);
    private sealed record NoticeNoteToken(int FirstPage, int EndPage);

    private sealed class NoticeNoteInventory {
        internal NoticeNoteInventory(IReadOnlyList<NoticeNoteGlyph> glyphs) {
            Glyphs = glyphs;
            Text = new string(glyphs.Select(glyph => glyph.Value).ToArray());
        }
        internal IReadOnlyList<NoticeNoteGlyph> Glyphs { get; }
        internal string Text { get; }
        internal NoticeNoteToken Token(string token) {
            Match match = Assert.Single(Regex.Matches(Text, Regex.Escape(PageFootnoteCompact(token))).Cast<Match>());
            return new NoticeNoteToken(Glyphs[match.Index].Page, Glyphs[match.Index + match.Length - 1].Page);
        }
        internal Letter[] PageLetters(int page) => Glyphs.Where(glyph => glyph.Page == page)
            .Select(glyph => glyph.Letter).Distinct().ToArray();
    }
}
