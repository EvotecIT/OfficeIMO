using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private const string TocFieldEndParagraphFixture = "word-generated-toc-field-end-paragraph.docx";
        private const string TocContentControlFixture = "word-generated-toc-gallery-pages.docx";

        [Theory]
        [InlineData(TocFieldEndParagraphFixture)]
        [InlineData(TocContentControlFixture)]
        public void Test_TocRegenerationKeepsEveryFieldEndAfterItsBegin(string fixture) {
            using WordDocument document = LoadTocFixture(fixture);

            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());

            AssertFieldCharactersNested(document._document.Body!);
            SdtContentBlock content = document.TableOfContent!.SdtBlock.SdtContentBlock!;
            // Word's ToC begins with its first entry (or the content control's title); there is no empty line above it.
            Paragraph first = content.Elements<Paragraph>().First(paragraph => !IsTocHeadingParagraph(paragraph));
            Assert.Contains(first.Descendants<FieldCode>(), code => code.Text.TrimStart().StartsWith("TOC", StringComparison.Ordinal));
            var errors = document.ValidateDocument();
            Assert.True(errors.Count == 0, FormatValidationErrors(errors));
        }

        [Fact]
        public void Test_TocWithoutTitleGainsNoTitleWhenRegenerated() {
            using WordDocument document = LoadTocFixture(TocFieldEndParagraphFixture);

            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());

            Paragraph first = document.TableOfContent!.SdtBlock.SdtContentBlock!.Elements<Paragraph>().First();
            Assert.StartsWith("Topic 1", string.Concat(first.Descendants<Text>().Select(text => text.Text)), StringComparison.Ordinal);
            Assert.Equal(30, document.TableOfContent.SdtBlock.Descendants<Hyperlink>().Count());
        }

        [Theory]
        [InlineData("figures")]
        [InlineData("index")]
        public void Test_CaptionListAndIndexRefreshKeepFieldCharactersNested(string kind) {
            using WordDocument document = LoadTocFixture(TocFieldEndParagraphFixture);
            WordTableOfContent table = document.TableOfContent!;

            if (kind == "figures") table.RefreshListOfFigures();
            else table.RefreshIndex();

            AssertFieldCharactersNested(document._document.Body!);
            var errors = document.ValidateDocument();
            Assert.True(errors.Count == 0, FormatValidationErrors(errors));
        }

        [Theory]
        [InlineData(TocFieldEndParagraphFixture)]
        [InlineData(TocContentControlFixture)]
        public void Test_WordTocFixturesKeepWordPagesInPdfAndFieldPagination(string fixture) {
            // Word lays out both documents on 6 pages: 28 ToC entries fit page 1, so Topic 1 starts on page 3.
            using (WordDocument document = LoadTocFixture(fixture)) {
                Assert.Equal(new TocPdfPages(6, 3, 3), ReadTocPdfPages(document));
            }

            using (WordDocument document = LoadTocFixture(fixture)) {
                document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
                Assert.Equal(new TocPdfPages(6, 3, 3), ReadTocPdfPages(document));
                // A second update starts from the content control OfficeIMO wrote.
                document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
                Assert.Equal(new TocPdfPages(6, 3, 3), ReadTocPdfPages(document));
            }
        }

        [Fact]
        public void Test_GeneratedTocUsesTheEntryStylesOfTheStoredToc() {
            using WordDocument generated = CreateLeveledHeadings();
            using WordDocument stored = CreateLeveledHeadings();
            stored.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
            stored.Settings.UpdateFieldsOnOpen = false;

            // The placeholder table is drawn from the headings; the regenerated one as stored. Both use TOC 1-3.
            IReadOnlyList<(string Text, double X, double Y)> generatedEntries = ReadEntryPositions(generated);
            IReadOnlyList<(string Text, double X, double Y)> storedEntries = ReadEntryPositions(stored);

            Assert.Equal(12, storedEntries.Count);
            Assert.Equal(storedEntries.Select(entry => entry.Text), generatedEntries.Select(entry => entry.Text));
            for (int index = 0; index < storedEntries.Count; index++) {
                Assert.Equal(storedEntries[index].X, generatedEntries[index].X, 1);
                Assert.Equal(storedEntries[index].Y, generatedEntries[index].Y, 1);
            }
        }

        private WordDocument LoadTocFixture(string fixture) {
            string path = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", fixture));
            using var stream = new MemoryStream(File.ReadAllBytes(path));
            return WordDocument.Load(stream);
        }

        private static WordDocument CreateLeveledHeadings() {
            WordDocument document = WordDocument.Create();
            document.AddTableOfContent();
            document.AddPageBreak();
            for (int index = 1; index <= 12; index++) {
                WordParagraphStyles style = (index % 3) switch {
                    1 => WordParagraphStyles.Heading1,
                    2 => WordParagraphStyles.Heading2,
                    _ => WordParagraphStyles.Heading3
                };
                document.AddParagraph("Entry " + index).Style = style;
                document.AddParagraph("Body " + index);
            }
            return document;
        }

        private static bool IsTocHeadingParagraph(Paragraph paragraph) =>
            !paragraph.Descendants<FieldChar>().Any() && !paragraph.Descendants<Hyperlink>().Any() && paragraph.InnerText.Length > 0;

        private static void AssertFieldCharactersNested(OpenXmlElement root) {
            int depth = 0;
            foreach (FieldChar fieldChar in root.Descendants<FieldChar>()) {
                if (fieldChar.FieldCharType?.Value == FieldCharValues.Begin) depth++;
                else if (fieldChar.FieldCharType?.Value == FieldCharValues.End) {
                    depth--;
                    Assert.True(depth >= 0, "A field end comes before its begin.");
                } else if (fieldChar.FieldCharType?.Value == FieldCharValues.Separate) {
                    Assert.True(depth > 0, "A field separator is outside a field.");
                }
            }
            Assert.Equal(0, depth);
        }

        private readonly record struct TocPdfPages(int PageCount, int TocTopicOnePage, int TopicOnePage);

        private static TocPdfPages ReadTocPdfPages(WordDocument document) {
            using var stream = new MemoryStream();
            document.SaveAsPdf(stream);
            IReadOnlyList<string> pages = PdfDocument.Load(stream.ToArray()).Reader.TextByPage();
            var entry = System.Text.RegularExpressions.Regex.Match(pages[0], @"Topic 1\b[\s.]*(\d+)");
            Assert.True(entry.Success, pages[0]);
            int topicOne = Enumerable.Range(0, pages.Count).First(index => pages[index].Contains("Body line for topic 1.", StringComparison.Ordinal)) + 1;
            return new TocPdfPages(pages.Count, int.Parse(entry.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture), topicOne);
        }

        private static IReadOnlyList<(string Text, double X, double Y)> ReadEntryPositions(WordDocument document) {
            using var stream = new MemoryStream();
            document.SaveAsPdf(stream);
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(stream.ToArray());
            return pdf.GetPage(1).GetWords()
                .Where(word => word.Text == "Entry")
                .Select(word => ("Entry", Math.Round(word.BoundingBox.Left, 1), Math.Round(word.BoundingBox.Bottom, 1)))
                .ToList();
        }
    }
}
