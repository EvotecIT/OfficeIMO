using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_LayoutPaginationMeasuresTextFieldsAfterTheirUpdate() {
            using WordDocument document = WordDocument.Create();
            document.SetDocumentVariable("Summary", string.Join(" ", Enumerable.Repeat("A long variable value that fills the page.", 120)));
            document.AddParagraph("Summary: ").AddField(WordFieldType.DocVariable, parameters: new List<string> { "Summary" });
            document.AddParagraph("Target").AddBookmark("AfterSummary");
            document.AddParagraph("See page ").AddField(WordFieldType.PageRef, parameters: new List<string> { "AfterSummary" });

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions().WithPdfLayoutPagination());

            string pageReference = report.Results.Single(x => x.FieldType == WordFieldType.PageRef).ResultText!;
            Assert.Equal(PageContaining(document, "Target"), int.Parse(pageReference, System.Globalization.CultureInfo.InvariantCulture));
            Assert.NotEqual("1", pageReference);
        }

        [Fact]
        public void Test_LayoutPaginationGivesAPageFieldThePageOfItsOwnLine() {
            using WordDocument document = WordDocument.Create();
            for (int index = 0; index < 10; index++) document.AddParagraph("Opening line " + index);
            WordParagraph paragraph = document.AddParagraph(string.Join(" ", Enumerable.Repeat("This paragraph starts on the first page and ends on the second.", 60)) + " Page ");
            paragraph.AddField(WordFieldType.Page);

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions().WithPdfLayoutPagination());

            Assert.Equal(1, PageContaining(document, "This paragraph starts"));
            Assert.Equal("2", report.Results.Single(x => x.FieldType == WordFieldType.Page).ResultText);
        }

        [Fact]
        public void Test_PdfExportDrawsStoredTocWithPageReferencesSplitAcrossRuns() {
            using WordDocument document = CreateTwoLevelChapters();
            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
            foreach (FieldCode code in document.TableOfContent!.SdtBlock.Descendants<FieldCode>().Where(code => code.Text.Contains("PAGEREF")).ToList()) {
                string text = code.Text;
                int split = text.IndexOf("REF", StringComparison.Ordinal);
                code.Text = text.Substring(0, split);
                code.Parent!.InsertAfterSelf(new Run(new FieldCode(text.Substring(split)) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
            }

            // Stored entries keep the TOC 2 indent (11 pt); a table generated from the headings would indent by 22 pt.
            Assert.Equal(11D, LevelTwoIndent(document), 1);
        }

        [Fact]
        public void Test_PdfExportLeavesARegeneratedTocUnchanged() {
            using WordDocument document = CreateTwoLevelChapters();
            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
            string before = document.TableOfContent!.SdtBlock.OuterXml;

            Assert.Equal(11D, LevelTwoIndent(document), 1);
            Assert.False(document.Settings.UpdateFieldsOnOpen);
            Assert.Equal(before, document.TableOfContent!.SdtBlock.OuterXml);
        }

        [Fact]
        public void Test_PdfExportRegistersFontsUsedOnlyByStoredTocEntries() {
            const string family = "OfficeIMO Toc Entry";
            using WordDocument document = CreateTwoLevelChapters();
            document.UpdateFields(new WordFieldUpdateOptions { UpdateTablesOfContents = true }.WithPdfLayoutPagination());
            foreach (Hyperlink link in document.TableOfContent!.SdtBlock.Descendants<Hyperlink>()) {
                foreach (Run run in link.Elements<Run>()) {
                    (run.RunProperties ??= new RunProperties()).RunFonts = new RunFonts { Ascii = family, HighAnsi = family };
                }
            }
            var pdf = new PdfCreateOptions();
            pdf.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family,
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, Enumerable.Range(32, 95).ToArray())));

            using var stream = new MemoryStream();
            document.SaveAsPdf(stream, new WordPdfExportOptions { PdfOptions = pdf });
            using var read = UglyToad.PdfPig.PdfDocument.Open(stream.ToArray());
            Assert.Contains(read.GetPage(1).Letters, letter => letter.FontName.Contains("OfficeIMOTocEntry", StringComparison.OrdinalIgnoreCase));
        }

        private static WordDocument CreateTwoLevelChapters() {
            WordDocument document = WordDocument.Create();
            document.AddTableOfContent();
            document.AddPageBreak();
            for (int index = 1; index <= 4; index++) {
                document.AddParagraph("Chapter " + index).Style = WordParagraphStyles.Heading1;
                document.AddParagraph("Section " + index).Style = WordParagraphStyles.Heading2;
            }
            return document;
        }

        private static double LevelTwoIndent(WordDocument document) {
            using var stream = new MemoryStream();
            document.SaveAsPdf(stream);
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(stream.ToArray());
            var words = pdf.GetPage(1).GetWords().ToList();
            double chapter = words.First(word => word.Text == "Chapter").BoundingBox.Left;
            double section = words.First(word => word.Text == "Section").BoundingBox.Left;
            return section - chapter;
        }

        private static int PageContaining(WordDocument document, string text) {
            using var stream = new MemoryStream();
            document.SaveAsPdf(stream);
            IReadOnlyList<string> pages = PdfDocument.Load(stream.ToArray()).Reader.TextByPage();
            return Enumerable.Range(0, pages.Count).First(index => pages[index].Contains(text, StringComparison.Ordinal)) + 1;
        }
    }
}
