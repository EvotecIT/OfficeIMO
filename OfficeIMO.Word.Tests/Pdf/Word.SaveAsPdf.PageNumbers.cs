using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_DefaultDoesNotInjectPageNumbers() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNoNumbers.docx");
            string pdfNoNumbers = Path.Combine(_directoryWithFiles, "PdfWithoutNumbers.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Page1");
                document.AddPageBreak();
                document.AddParagraph("Page2");
                document.Save();
                document.SaveAsPdf(pdfNoNumbers);
            }

            Assert.True(File.Exists(pdfNoNumbers));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfNoNumbers)) {
                Assert.DoesNotContain("1/2", pdf.GetPage(1).Text);
                Assert.DoesNotContain("2/2", pdf.GetPage(2).Text);
            }
        }

        [Fact]
        public void SaveAsPdf_Formats_PageNumbers() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNumbersFormat.docx");
            string pdfCustom = Path.Combine(_directoryWithFiles, "PdfCustomNumbers.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Page1");
                document.AddPageBreak();
                document.AddParagraph("Page2");
                document.Save();
                document.SaveAsPdf(pdfCustom, new PdfSaveOptions {
                    IncludePageNumbers = true,
                    PageNumberFormat = "Page {current} of {total}"
                });
            }

            Assert.True(File.Exists(pdfCustom));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfCustom)) {
                Assert.Contains("Page 1 of 2", pdf.GetPage(1).Text);
                Assert.Contains("Page 2 of 2", pdf.GetPage(2).Text);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_PageBreaks() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativePageBreaks.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativePageBreaks.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native page break");
                document.AddPageBreak();
                document.AddParagraph("After native page break");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                Assert.Equal(2, pdf.NumberOfPages);
                Assert.Contains("Before native page break", pdf.GetPage(1).Text);
                Assert.Contains("After native page break", pdf.GetPage(2).Text);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Word_Section_PageNumbering() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionPageNumbering.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionPageNumbering.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Sections[0].AddPageNumbering(3, NumberFormatValues.UpperRoman);
                document.AddParagraph("Native section page numbering first page");
                document.AddPageBreak();
                document.AddParagraph("Native section page numbering second page");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = true
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                Assert.Equal(2, pdf.NumberOfPages);

                string page1Text = NormalizeNativePageNumberText(pdf.GetPage(1).Text);
                string page2Text = NormalizeNativePageNumberText(pdf.GetPage(2).Text);

                Assert.Contains("Nativesectionpagenumberingfirstpage", page1Text, StringComparison.Ordinal);
                Assert.Contains("Nativesectionpagenumberingsecondpage", page2Text, StringComparison.Ordinal);
                Assert.Contains("III/IV", page1Text, StringComparison.Ordinal);
                Assert.Contains("IV/IV", page2Text, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Keeps_NumPages_As_Document_Total_When_Section_Restarts() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeNumPagesDocumentTotal.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeNumPagesDocumentTotal.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddHeadersAndFooters();
                RequireSectionFooter(document, 0, HeaderFooterValues.Default)
                    .AddParagraph("First field footer ")
                    .AddPageNumber(includeTotalPages: true, separator: " / ");
                document.AddParagraph("First section first page");
                document.AddPageBreak();
                document.AddParagraph("First section second page");

                WordSection secondSection = document.AddSection();
                secondSection.AddPageNumbering(1, NumberFormatValues.Decimal);
                RequireSectionFooter(document, 1, HeaderFooterValues.Default)
                    .AddParagraph("Restart field footer ")
                    .AddPageNumber(includeTotalPages: true, separator: " / ");
                secondSection.AddParagraph("Restarted section first page");
                document.AddPageBreak();
                secondSection.AddParagraph("Restarted section second page");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            Assert.Equal(4, pdf.NumberOfPages);
            string page3Text = NormalizeNativePageNumberText(pdf.GetPage(3).Text);
            string page4Text = NormalizeNativePageNumberText(pdf.GetPage(4).Text);
            Assert.Contains("Restartfieldfooter1/4", page3Text, StringComparison.Ordinal);
            Assert.Contains("Restartfieldfooter2/4", page4Text, StringComparison.Ordinal);
            Assert.DoesNotContain("Restartfieldfooter1/2", page3Text, StringComparison.Ordinal);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_SectionPages_Field_To_Section_Total() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionPagesField.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeSectionPagesField.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddHeadersAndFooters();
                WordParagraph firstFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("First section ");
                firstFooter.AddField(WordFieldType.Page);
                firstFooter.AddText(" / ");
                firstFooter.AddField(WordFieldType.SectionPages);
                firstFooter.AddText(" / ");
                firstFooter.AddField(WordFieldType.NumPages);
                document.AddParagraph("First section first page");
                document.AddPageBreak();
                document.AddParagraph("First section second page");

                WordSection secondSection = document.AddSection();
                secondSection.AddPageNumbering(1, NumberFormatValues.Decimal);
                WordParagraph secondFooter = RequireSectionFooter(document, 1, HeaderFooterValues.Default).AddParagraph("Second section ");
                secondFooter.AddField(WordFieldType.Page);
                secondFooter.AddText(" / ");
                secondFooter.AddField(WordFieldType.SectionPages);
                secondFooter.AddText(" / ");
                secondFooter.AddField(WordFieldType.NumPages);
                secondSection.AddParagraph("Second section first page");
                document.AddPageBreak();
                secondSection.AddParagraph("Second section second page");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            Assert.Equal(4, pdf.NumberOfPages);
            string page1Text = NormalizeNativePageNumberText(pdf.GetPage(1).Text);
            string page3Text = NormalizeNativePageNumberText(pdf.GetPage(3).Text);

            Assert.Contains("Firstsection1/2/4", page1Text, StringComparison.Ordinal);
            Assert.Contains("Secondsection1/2/4", page3Text, StringComparison.Ordinal);
            Assert.DoesNotContain("Secondsection1/4/4", page3Text, StringComparison.Ordinal);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_PageFields_To_PageTokens() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterPageFields.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterPageFields.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddHeadersAndFooters();
                WordParagraph footerParagraph = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("Native field footer ");
                footerParagraph.AddPageNumber(includeTotalPages: true, separator: " / ");
                document.AddParagraph("Native field footer first page");
                document.AddPageBreak();
                document.AddParagraph("Native field footer second page");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    PageNumberFormat = "AUTO {current}/{total}"
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                Assert.Equal(2, pdf.NumberOfPages);

                string page1Text = NormalizeNativePageNumberText(pdf.GetPage(1).Text);
                string page2Text = NormalizeNativePageNumberText(pdf.GetPage(2).Text);

                Assert.Contains("Nativefieldfooter1/2", page1Text, StringComparison.Ordinal);
                Assert.Contains("Nativefieldfooter2/2", page2Text, StringComparison.Ordinal);
                Assert.DoesNotContain("AUTO", page1Text, StringComparison.Ordinal);
                Assert.DoesNotContain("AUTO", page2Text, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_HeaderFooter_PageField_Formats_To_PageTokenStyle() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterPageFieldFormats.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterPageFieldFormats.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddHeadersAndFooters();
                WordParagraph footerParagraph = RequireSectionFooter(document, 0, HeaderFooterValues.Default).AddParagraph("Native roman field footer ");
                footerParagraph.AddPageNumber(includeTotalPages: true, format: WordFieldFormat.Roman, separator: " / ");
                document.AddParagraph("Native roman field footer first page");
                document.AddPageBreak();
                document.AddParagraph("Native roman field footer second page");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions());
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                Assert.Equal(2, pdf.NumberOfPages);

                string page1Text = NormalizeNativePageNumberText(pdf.GetPage(1).Text);
                string page2Text = NormalizeNativePageNumberText(pdf.GetPage(2).Text);

                Assert.Contains("NativeromanfieldfooterI/II", page1Text, StringComparison.Ordinal);
                Assert.Contains("NativeromanfieldfooterII/II", page2Text, StringComparison.Ordinal);
            }
        }

        private static string NormalizeNativePageNumberText(string text) =>
            string.Concat(text.Where(c => !char.IsWhiteSpace(c)));
    }
}
