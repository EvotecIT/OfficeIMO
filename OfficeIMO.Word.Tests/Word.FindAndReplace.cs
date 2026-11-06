using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DocumentFindAndReplace() {
            string filePath = Path.Combine(_directoryWithFiles, "SimpleWordDocumentSearchFunctionality.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test Section");

                document.Paragraphs[0].AddComment("Przemysław", "PK", "This is my comment");

                document.AddParagraph("Test Section - another line");

                document.Paragraphs[1].AddComment("Przemysław", "PK", "More comments");

                document.AddParagraph("This is a text ").AddText("more text").AddText(" even longer text").AddText(" and Even longer right?");

                document.AddParagraph("this is a text ").AddText("more text 1").AddText(" even longer text 1").AddText(" and Even longer right?");

                // we now ensure that we add bold to complicate the search
                document.Paragraphs[9].Bold = true;
                document.Paragraphs[10].Bold = true;

                var listFound = document.Find("Test Section");
                Assert.Equal(2, listFound.MatchCount);
                Assert.Equal(2, listFound.BodyParagraphs.Count);

                var replacedCount = document.ReplaceAll("Test Section", "Production Section");
                Assert.True(replacedCount == 2);

                // should be 0 because it stretches over 2 paragraphs
                var replacedCount1 = document.ReplaceAll("This is a text more text", "Shorter text");
                Assert.True(replacedCount1 == 2);

                document.CleanupDocument();

                // cleanup should merge paragraphs making it easier to find and replace text
                // this only works for same formatting though
                // may require improvement in the future to ignore formatting completely, but then it's a bit tricky which formatting to apply
                var replacedCount2 = document.ReplaceAll("This is a text more text", "Shorter text");
                Assert.True(replacedCount2 == 0);

                var replacedCount3 = document.ReplaceAll("even longer", "not longer");
                Assert.True(replacedCount3 == 4);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentSearchFunctionality.docx"))) {

                Assert.True(document.Paragraphs[0].Text == "Production Section");

                var table = document.AddTable(3, 3);
                table.Rows[0].Cells[0].Paragraphs[0].AddText("Test Section");
                table.Rows[0].Cells[1].Paragraphs[0].AddText("Test Section");
                table.Rows[0].Cells[2].Paragraphs[0].AddText("Test ").AddText("Sect").AddText("ion");

                _ = document.PageSetup.Header;

                var header = document.PageSetup.GetHeader(WordHeaderFooterType.Default);
                var tableInHeader = document.AddTable(3, 3);
                tableInHeader.Rows[0].Cells[0].Paragraphs[0].AddText("Prod Section");
                tableInHeader.Rows[0].Cells[1].Paragraphs[0].AddText("Prod Section");
                tableInHeader.Rows[0].Cells[2].Paragraphs[0].AddText("Prod ").AddText("Sect").AddText("ion");

                var footer = document.PageSetup.GetFooter(WordHeaderFooterType.Default);
                var tableInFooter = document.AddTable(3, 3);
                tableInFooter.Rows[0].Cells[0].Paragraphs[0].AddText("Prod Section");
                tableInFooter.Rows[0].Cells[1].Paragraphs[0].AddText("Prod Section");
                tableInFooter.Rows[0].Cells[2].Paragraphs[0].AddText("Prod ").AddText("Sect").AddText("ion");

                document.CleanupDocument();

                var listFound1 = document.Find("Test Section");
                Assert.True(listFound1.Paragraphs.Count == 3);

                var listFound2 = document.Find("Prod Section");
                Assert.True(listFound2.Paragraphs.Count == 6);

                var replacedCount = document.ReplaceAll("Prod Section", "Production Section");
                Assert.True(replacedCount == 6);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentSearchFunctionality.docx"))) {



                document.Save();
            }
        }

        [Fact]
        public void Test_DocumentRegexFind() {
            string filePath = Path.Combine(_directoryWithFiles, "RegexWordDocument.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Item 123");
                document.AddParagraph("Item 456");
                document.AddParagraph("Different 789");

                var regex = new Regex(@"Item \d{3}");
                var result = document.Find(regex);

                Assert.Equal(2, result.MatchCount);
                Assert.Equal(2, result.BodyParagraphs.Count);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var regex = new Regex(@"Item \d{3}");
                var result = document.Find(regex);

                Assert.Equal(2, result.MatchCount);
                Assert.Equal(2, result.BodyParagraphs.Count);
            }
        }

        [Fact]
        public void DocumentSearchAndReplace_VisitsLaterAndSharedHeaderFooterStoriesOnce() {
            using var source = WordDocument.Create();
            var later = source.AddSection();
            var shared = source.AddSection();
            source.Sections[0].PageSetup.Header.AddParagraph("Unrelated first header");
            later.PageSetup.Header.AddParagraph("TOKEN later header");
            later.PageSetup.Footer.AddParagraph("TOKEN later footer");
            // A second explicit reference to the same part must not count as a second story.
            shared._sectionProperties.AddChild(later._sectionProperties.Elements<HeaderReference>().Single().CloneNode(true), true);

            using var input = new MemoryStream();
            source.Save(input);
            input.Position = 0;
            using var document = WordDocument.Load(input, new WordLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit });

            foreach (WordSearchResult result in new[] { document.Find("TOKEN"), document.Find(new Regex("TOKEN")) }) {
                Assert.Equal(2, result.MatchCount);
                Assert.Equal("TOKEN later header", Assert.Single(result.DefaultHeaderParagraphs).Text);
                Assert.Equal("TOKEN later footer", Assert.Single(result.DefaultFooterParagraphs).Text);
            }
            Assert.Equal(2, document.Sections[2].Find("TOKEN").MatchCount);
            // Keep the original token in the replacement so a duplicate edit is observable.
            Assert.Equal(2, document.ReplaceAll("TOKEN", "TOKEN filled"));

            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;
            using var reopened = WordDocument.Load(output, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.Equal(new[] { "TOKEN filled later header", "TOKEN filled later footer" },
                reopened.Find("TOKEN").Paragraphs.Select(paragraph => paragraph.Text));
        }

        [Fact]
        public void DocumentSearchAndReplace_VisitsOrdinaryNotesAsSeparateStories() {
            using var source = WordDocument.Create();
            source.AddParagraph("First reference").AddFootNote("TOKEN alpha");
            source.AddParagraph("Second reference").AddFootNote("beta TOKEN");
            source.AddParagraph("Endnote reference").AddEndNote("TOKEN endnote");
            var separator = source._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
                .Elements<Footnote>().First(note => note.Type?.Value == FootnoteEndnoteValues.Separator);
            separator.Elements<Paragraph>().Single().Append(new Run(new Text("TOKEN separator")));

            using var input = new MemoryStream();
            source.Save(input);
            input.Position = 0;
            using var document = WordDocument.Load(input, new WordLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit });

            foreach (WordSearchResult result in new[] { document.Find("TOKEN"), document.Find(new Regex("TOKEN")) }) {
                Assert.Equal(3, result.MatchCount);
                Assert.Equal(new[] { "TOKEN alpha", "beta TOKEN", "TOKEN endnote" },
                    result.Paragraphs.Select(paragraph => paragraph.Text));
            }
            Assert.Equal(0, document.ReplaceAll("alphabeta", "unexpected"));
            Assert.Equal(3, document.ReplaceAll("TOKEN", "filled"));

            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;
            using var reopened = WordDocument.Load(output, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.Equal(new[] { "filled alpha", "beta filled", "filled endnote" },
                reopened.Find("filled").Paragraphs.Select(paragraph => paragraph.Text));
            Assert.Contains("TOKEN separator", reopened._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.InnerText);
        }

        [Fact]
        public void Test_FindAndReplace_WithNewLines() {
            string filePath = Path.Combine(_directoryWithFiles, "FindAndReplaceNewLines.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Before KEY After");
                paragraph.SetBold();

                var replaced = document.ReplaceAll("KEY", "Line1\nLine2");

                Assert.Equal(1, replaced);
                Assert.Equal("Before Line1\nLine2 After", paragraph.Text);

                var run = paragraph.GetRuns().First();
                Assert.NotNull(run.Break);
                Assert.True(paragraph.Bold);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var paragraph = document.Paragraphs[0];
                Assert.Equal("Before Line1\nLine2 After", paragraph.Text);

                var run = paragraph.GetRuns().First();
                Assert.NotNull(run.Break);
                Assert.True(paragraph.Bold);
            }
        }

        [Fact]
        public void FindAndReplace_PreservesRunsOnContainedMatches() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("Before ");
            paragraph.AddText("KEY").SetBold();
            paragraph.AddText(" after");

            int replaced = document.ReplaceAll("KEY", "VALUE", new OfficeTextSearchOptions { MatchCase = true });

            Assert.Equal(1, replaced);
            Assert.Equal("Before VALUE after", string.Concat(paragraph._paragraph.Descendants<Text>().Select(text => text.Text)));
            var runs = paragraph._paragraph.Elements<Run>().ToList();
            Assert.Equal(3, runs.Count);
            Assert.NotNull(runs[1].RunProperties?.Bold);
        }

        [Fact]
        public void FindAndReplace_RetainsCrossParagraphFallback() {
            using WordDocument document = WordDocument.Create();
            WordParagraph first = document.AddParagraph("Alpha");
            WordParagraph second = document.AddParagraph("Beta");

            int replaced = document.ReplaceAll("AlphaBeta", "Joined", new OfficeTextSearchOptions { MatchCase = true });

            Assert.Equal(1, replaced);
            Assert.Equal("Joined", first.Text);
            Assert.Equal(string.Empty, second.Text);
        }

        [Fact]
        public void FindAndReplace_RetainsOrdinalFallbackAcrossThreeRuns() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("Alpha");
            paragraph.AddText(string.Empty);
            paragraph.AddText("Beta");

            int replaced = document.ReplaceAll("AlphaBeta", "Joined", new OfficeTextSearchOptions { MatchCase = true });

            Assert.Equal(1, replaced);
            Assert.Equal("Joined", paragraph.Text);
        }
    }
}
