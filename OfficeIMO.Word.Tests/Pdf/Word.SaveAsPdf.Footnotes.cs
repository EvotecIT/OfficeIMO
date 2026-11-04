using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Linq;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
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
                ? document.AddTable(1, 1, WordTableStyle.TableGrid).Rows[0].Cells[0].Paragraphs[0].SetText("SameNumberMarker")
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
            using var imported = nativeDoc ? WordDocument.Load(new MemoryStream(document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }))) : null;
            (imported ?? document).SaveAsPdf(path, new WordPdfExportOptions { IncludePageNumbers = false });
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
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Native footnote here1", allText);
                AssertPdfFootnoteAssociation(pdf, "Native footnote here", "Native footnote text", "1");
                Assert.Contains("Native after footnote", allText);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Endnote_Markers_And_Text() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeEndnotes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeEndnotes.pdf");
            string destinationName;

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph footnoteParagraph = document.AddParagraph("Native footnote here");
                footnoteParagraph.AddFootNote("Native footnote text");
                WordParagraph endnoteParagraph = document.AddParagraph("Native endnote here");
                endnoteParagraph.AddEndNote("Native endnote text");
                document.AddParagraph("Native after notes");
                document.Save();
                long noteId = Assert.Single(endnoteParagraph._paragraph.Descendants<W.EndnoteReference>()).Id!.Value;
                destinationName = "OfficeIMO.Endnote." + noteId;
                var main = document._wordprocessingDocument.MainDocumentPart!;
                string documentBefore = main.Document.OuterXml;
                string notesBefore = main.EndnotesPart!.Endnotes!.OuterXml;
                byte[] sourceBefore = File.ReadAllBytes(docPath);
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
                Assert.Equal(documentBefore, main.Document.OuterXml);
                Assert.Equal(notesBefore, main.EndnotesPart.Endnotes.OuterXml);
                Assert.Equal(sourceBefore, File.ReadAllBytes(docPath));
            }

            Assert.True(File.Exists(pdfPath));
            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Native footnote here1", allText);
                Assert.Contains("Native endnote herei", allText);
                AssertPdfFootnoteAssociation(pdf, "Native footnote here", "Native footnote text", "1");
                // Keep the supplied phrase's authored spaces; no space is
                // authored between its separate endnote marker and the text.
                Assert.Contains("Native endnote text", allText);
                Assert.Contains("Native after notes", allText);
                var notePage = Assert.Single(pdf.GetPages(), page => page.Text.Contains("Native endnote text", StringComparison.Ordinal));
                var noteText = AnchoredRunningLetters(notePage, "Native endnote text");
                var label = Assert.Single(notePage.Letters, letter => letter.Value == "i" &&
                    letter.StartBaseLine.Y > noteText[0].StartBaseLine.Y &&
                    letter.StartBaseLine.Y - noteText[0].StartBaseLine.Y < noteText[0].PointSize &&
                    Math.Abs(letter.EndBaseLine.X - noteText[0].StartBaseLine.X) <= .15D);
                Assert.True(label.PointSize < noteText[0].PointSize);
                var followingPage = Assert.Single(pdf.GetPages(), page => page.Text.Contains("Native after notes", StringComparison.Ordinal));
                Assert.True(followingPage.Number <= notePage.Number);
                if (followingPage.Number == notePage.Number)
                    Assert.True(AnchoredRunningLetters(followingPage, "Native after notes")[0].StartBaseLine.Y > label.StartBaseLine.Y);
                PdfDocumentInfo navigation = PdfInspector.Inspect(File.ReadAllBytes(pdfPath));
                PdfLinkAnnotation annotation = Assert.Single(navigation.LinkAnnotations,
                    link => link.DestinationName == destinationName);
                Assert.True(annotation.IsNamedDestinationLink);
                Assert.True(annotation.Width > 0D && annotation.Height > 0D);
                Assert.Equal(Assert.Single(pdf.GetPages(), page => page.Text.Contains("Native endnote here", StringComparison.Ordinal)).Number,
                    annotation.PageNumber);
                PdfNamedDestination destination = Assert.Single(navigation.NamedDestinations,
                    item => item.Name == destinationName);
                Assert.Equal(notePage.Number, destination.PageNumber);
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
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions {
                    IncludePageNumbers = false
                });
            }

            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("First section note1", allText);
                Assert.Contains("Second section note2", allText);
                var first = AssertPdfFootnoteAssociation(pdf, "First section note", "First section footnote", "1");
                var second = AssertPdfFootnoteAssociation(pdf, "Second section note", "Second section footnote", "2");
                AssertPdfFootnoteSourceOrder(first, second);
                Assert.DoesNotContain("Second section note1", allText);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Numbers_Nested_Table_Footnotes_In_Document_Order() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeNestedTableFootnotes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeNestedTableFootnotes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable outer = document.AddTable(1, 2, WordTableStyle.TableGrid);
                WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 1, WordTableStyle.TableGrid);
                nested.Rows[0].Cells[0].Paragraphs[0]
                    .SetText("Nested table note")
                    .AddFootNote("Nested table footnote");
                outer.Rows[0].Cells[1].Paragraphs[0]
                    .SetText("Later outer note")
                    .AddFootNote("Later outer footnote");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions { IncludePageNumbers = false });
            }

            using (var pdf = PdfPigDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Later outer note2", allText);
                var nested = AssertPdfFootnoteAssociation(pdf, "Nested table note", "Nested table footnote", "1");
                var outer = AssertPdfFootnoteAssociation(pdf, "Later outer note", "Later outer footnote", "2");
                // Side-by-side cells can place a nested reference lower than a later source reference.
                // The labels and footnote bodies retain source order regardless of those baselines.
                Assert.True(PdfFootnoteGlyphPrecedes(nested.NotePage, nested.NoteLabel, outer.NotePage, outer.NoteLabel),
                    "Footnote bodies must retain their source order.");
                Assert.DoesNotContain("Later outer note1", allText);
            }
        }
    }
}
