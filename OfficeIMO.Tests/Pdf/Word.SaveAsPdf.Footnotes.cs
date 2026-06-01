using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
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
            using (var pdf = PdfDocument.Open(pdfPath)) {
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
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (var pdf = PdfDocument.Open(pdfPath)) {
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
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (var pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                string normalizedText = Regex.Replace(allText, @"\s+", " ");
                Assert.Contains("Native footnote here1", allText);
                Assert.Contains("Native endnote here1", allText);
                Assert.Contains("1 Native footnote text", normalizedText);
                Assert.Contains("1 Native endnote text", normalizedText);
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
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using (var pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                string normalizedText = Regex.Replace(allText, @"\s+", " ");
                Assert.Contains("First section note1", allText);
                Assert.Contains("Second section note2", allText);
                Assert.Contains("1 First section footnote", normalizedText);
                Assert.Contains("2 Second section footnote", normalizedText);
                Assert.DoesNotContain("Second section note1", allText);
            }
        }
    }
}
