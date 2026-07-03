using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocHyperlinkDisplayBookmarksAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph bodyParagraph = document.AddParagraph("Body ");
                    bodyParagraph.AddHyperLink("placeholder", new Uri("https://officeimo.net/body-bookmark"), addStyle: true);
                    ReplaceHyperlinkDisplayWithBookmark(bodyParagraph.Hyperlink!._hyperlink, "71", "BodyHyperlinkBookmark", "BodyMarked");

                    WordTable table = document.AddTable(1, 1);
                    WordParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph("Cell ", removeExistingParagraphs: true);
                    cellParagraph.AddHyperLink("placeholder", new Uri("mailto:cell-bookmark@example.org"), addStyle: true);
                    ReplaceHyperlinkDisplayWithBookmark(cellParagraph.Hyperlink!._hyperlink, "72", "CellHyperlinkBookmark", "CellMarked");

                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    WordParagraph headerParagraph = section.Header.Default!.AddParagraph("Header ");
                    headerParagraph.AddHyperLink("placeholder", new Uri("https://officeimo.net/header-bookmark"), addStyle: true);
                    ReplaceHyperlinkDisplayWithBookmark(headerParagraph.Hyperlink!._hyperlink, "73", "HeaderHyperlinkBookmark", "HeaderMarked");

                    WordParagraph footerParagraph = section.Footer.Default!.AddParagraph("Footer ");
                    footerParagraph.AddHyperLink("placeholder", new Uri("mailto:footer-bookmark@example.org"), addStyle: true);
                    ReplaceHyperlinkDisplayWithBookmark(footerParagraph.Hyperlink!._hyperlink, "74", "FooterHyperlinkBookmark", "FooterMarked");

                    WordParagraph noteReferences = document.AddParagraph("Notes ");
                    WordParagraph footnoteReference = noteReferences.AddFootNote("footnote placeholder");
                    WordParagraph footnoteBody = footnoteReference.FootNote!.Paragraphs![1];
                    footnoteBody.AddHyperLink("placeholder", new Uri("https://officeimo.net/footnote-bookmark"), addStyle: true);
                    ReplaceHyperlinkDisplayWithBookmark(footnoteBody.Hyperlink!._hyperlink, "75", "FootnoteHyperlinkBookmark", "FootnoteMarked");

                    WordParagraph endnoteReference = noteReferences.AddEndNote("endnote placeholder");
                    WordParagraph endnoteBody = endnoteReference.EndNote!.Paragraphs![1];
                    endnoteBody.AddHyperLink("placeholder", new Uri("mailto:endnote-bookmark@example.org"), addStyle: true);
                    ReplaceHyperlinkDisplayWithBookmark(endnoteBody.Hyperlink!._hyperlink, "76", "EndnoteHyperlinkBookmark", "EndnoteMarked");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                AssertBookmarkRoundTrip(reloaded, "BodyHyperlinkBookmark");
                AssertBookmarkRoundTrip(reloaded, "CellHyperlinkBookmark");
                AssertBookmarkRoundTrip(reloaded, "HeaderHyperlinkBookmark");
                AssertBookmarkRoundTrip(reloaded, "FooterHyperlinkBookmark");
                AssertBookmarkRoundTrip(reloaded, "FootnoteHyperlinkBookmark");
                AssertBookmarkRoundTrip(reloaded, "EndnoteHyperlinkBookmark");

                Assert.Contains(DistinctHyperlinks(reloaded.HyperLinks), link => GetHyperlinkText(link._hyperlink) == "BodyMarked");
                Assert.Contains(
                    DistinctHyperlinks(Assert.Single(reloaded.Tables).Rows[0].Cells[0].Paragraphs.SelectMany(paragraph => paragraph.GetRuns()).Where(run => run.IsHyperLink).Select(run => run.Hyperlink)),
                    link => GetHyperlinkText(link!._hyperlink) == "CellMarked");

                WordSection reloadedSection = Assert.Single(reloaded.Sections);
                Assert.Contains(
                    DistinctHyperlinks(reloadedSection.Header.Default!.Paragraphs.SelectMany(paragraph => paragraph.GetRuns()).Where(run => run.IsHyperLink).Select(run => run.Hyperlink)),
                    link => GetHyperlinkText(link!._hyperlink) == "HeaderMarked");
                Assert.Contains(
                    DistinctHyperlinks(reloadedSection.Footer.Default!.Paragraphs.SelectMany(paragraph => paragraph.GetRuns()).Where(run => run.IsHyperLink).Select(run => run.Hyperlink)),
                    link => GetHyperlinkText(link!._hyperlink) == "FooterMarked");

                WordFootNote footnote = Assert.Single(reloaded.FootNotes);
                Assert.Contains(
                    DistinctHyperlinks(footnote.Paragraphs!.Where(run => run.IsHyperLink).Select(run => run.Hyperlink)),
                    link => GetHyperlinkText(link!._hyperlink) == "FootnoteMarked");

                WordEndNote endnote = Assert.Single(reloaded.EndNotes);
                Assert.Contains(
                    DistinctHyperlinks(endnote.Paragraphs!.Where(run => run.IsHyperLink).Select(run => run.Hyperlink)),
                    link => GetHyperlinkText(link!._hyperlink) == "EndnoteMarked");
            } finally {
                DeleteIfExists(docPath);
            }
        }

        private static void ReplaceHyperlinkDisplayWithBookmark(Hyperlink hyperlink, string id, string name, string text) {
            hyperlink.RemoveAllChildren();
            hyperlink.Append(
                new BookmarkStart { Id = id, Name = name },
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
                new BookmarkEnd { Id = id });
        }

        private static void AssertBookmarkRoundTrip(WordDocument document, string bookmarkName) {
            Assert.Contains(document.Bookmarks, bookmark => bookmark.Name == bookmarkName);
        }
    }
}
