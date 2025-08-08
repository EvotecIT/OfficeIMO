using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_Can_Disable_PageNumbers() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNoNumbers.docx");
            string pdfNoNumbers = Path.Combine(_directoryWithFiles, "PdfWithoutNumbers.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Page1");
                document.AddPageBreak();
                document.AddParagraph("Page2");
                document.Save();
                document.SaveAsPdf(pdfNoNumbers, new PdfSaveOptions { IncludePageNumbers = false });
            }

            Assert.True(File.Exists(pdfNoNumbers));
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
                document.SaveAsPdf(pdfCustom, new PdfSaveOptions { PageNumberFormat = "Page {current} of {total}" });
            }

            Assert.True(File.Exists(pdfCustom));
        }
    }
}
