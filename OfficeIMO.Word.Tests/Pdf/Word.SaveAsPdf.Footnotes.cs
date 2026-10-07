using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData("body", false)]
        [InlineData("table", false)]
        [InlineData("control", false)]
        [InlineData("heading", false)]
        [InlineData("body", true)]
        [InlineData("table", true)]
        [InlineData("control", true)]
        [InlineData("heading", true)]
        public void SaveAsPdf_PreservesFootnoteAndEndnoteWithTheSameDisplayNumber(string context, bool nativeDoc) {
            string path = Path.Combine(_directoryWithFiles, "SameNumberNotes" + context + ".pdf");
            using var document = WordDocument.Create();
            document.Sections[0].AddEndnoteProperties(WordNumberFormat.Decimal);
            WordParagraph paragraph = context == "table"
                ? document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].SetText("SameNumberMarker")
                : document.AddParagraph("SameNumberMarker");
            paragraph.AddFootNote("FootnoteBody");
            paragraph.AddEndNote("EndnoteBody");
            if (context == "heading") paragraph.Style = WordParagraphStyles.Heading7;
            foreach (var run in paragraph._paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()) {
                int size = run.Elements<DocumentFormat.OpenXml.Wordprocessing.FootnoteReference>().Any() ? 14 :
                    run.Elements<DocumentFormat.OpenXml.Wordprocessing.EndnoteReference>().Any() ? 16 : 0;
                if (size == 0) continue;
                run.RunProperties ??= new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
                run.RunProperties.FontSize = new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = (size * 2).ToString() };
            }
            if (context == "control") {
                var content = new DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock(paragraph._paragraph!.CloneNode(true));
                paragraph._paragraph.InsertBeforeSelf(new DocumentFormat.OpenXml.Wordprocessing.SdtBlock(content));
                paragraph._paragraph.Remove();
            }
            using var imported = nativeDoc ? WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc))) : null;
            (imported ?? document).SaveAsPdf(path, new WordToPdfOptions { IncludePageNumbers = false });
            var spans = OfficeIMO.Pdf.PdfReadDocument.Open(File.ReadAllBytes(path)).Pages.SelectMany(page => page.GetTextSpans()).ToArray();
            var marker = Assert.Single(spans, span => span.Text.Contains("SameNumberMarker"));
            Assert.Equal(2, spans.Count(span => span.Text == "1" && span.Y > marker.Y && span.Y - marker.Y < marker.FontSize));
            var references = spans.Where(span => span.Text == "1" && span.Y > marker.Y && span.Y - marker.Y < marker.FontSize)
                .OrderBy(span => span.FontSize).ToArray();
            Assert.Equal(14D * .65D, references[0].FontSize, precision: 3);
            Assert.Equal(16D * .65D, references[1].FontSize, precision: 3);
        }

        [Fact]
        public void SaveAsPdf_Renders_Footnotes_And_PageNumbers() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfFootnotes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfFootnotes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph p = document.AddParagraph("Footnote here");
                p.AddFootNote("Footnote text");
                document.Save();
                document.SaveAsPdf(pdfPath);
            }

            Assert.True(File.Exists(pdfPath));
            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Footnote here1", allText);
                Assert.Equal(1, pdf.NumberOfPages);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Footnote_Markers_And_Text() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeFootnotes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeFootnotes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph first = document.AddParagraph("Native footnote here");
                first.AddFootNote("Native footnote text");
                document.AddParagraph("Native after footnote");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Native footnote here1", allText);
                Assert.Contains("1 Native footnote text", Regex.Replace(allText, @"\s+", " "));
                Assert.Contains("Native after footnote", allText);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Endnote_Markers_And_Text() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeEndnotes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeEndnotes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph footnoteParagraph = document.AddParagraph("Native footnote here");
                footnoteParagraph.AddFootNote("Native footnote text");
                WordParagraph endnoteParagraph = document.AddParagraph("Native endnote here");
                endnoteParagraph.AddEndNote("Native endnote text");
                document.AddParagraph("Native after notes");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                string normalizedText = Regex.Replace(allText, @"\s+", " ");
                Assert.Contains("Native footnote here1", allText);
                Assert.Contains("Native endnote herei", allText);
                Assert.Contains("1 Native footnote text", normalizedText);
                Assert.Contains("i Native endnote text", normalizedText);
                Assert.Contains("Native after notes", allText);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Keeps_Footnote_Numbering_Continuous_Across_Sections() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionFootnotes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionFootnotes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("First section note").AddFootNote("First section footnote");
                WordSection secondSection = document.AddSection();
                secondSection.AddParagraph("Second section note").AddFootNote("Second section footnote");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                    IncludePageNumbers = false
                });
            }

            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                string normalizedText = Regex.Replace(allText, @"\s+", " ");
                Assert.Contains("First section note1", allText);
                Assert.Contains("Second section note2", allText);
                Assert.Contains("1 First section footnote", normalizedText);
                Assert.Contains("2 Second section footnote", normalizedText);
                Assert.DoesNotContain("Second section note1", allText);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Numbers_Nested_Table_Footnotes_In_Document_Order() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeNestedTableFootnotes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeNestedTableFootnotes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable outer = document.AddTable(1, 2);
                WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 1);
                nested.Rows[0].Cells[0].Paragraphs[0]
                    .SetText("Nested table note")
                    .AddFootNote("Nested table footnote");
                outer.Rows[0].Cells[1].Paragraphs[0]
                    .SetText("Later outer note")
                    .AddFootNote("Later outer footnote");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
            }

            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                string normalizedText = Regex.Replace(allText, @"\s+", " ");
                Assert.Contains("Later outer note2", allText);
                Assert.Contains("1 Nested table footnote", normalizedText);
                Assert.Contains("2 Later outer footnote", normalizedText);
                Assert.DoesNotContain("Later outer note1", allText);
            }
        }
    }
}
