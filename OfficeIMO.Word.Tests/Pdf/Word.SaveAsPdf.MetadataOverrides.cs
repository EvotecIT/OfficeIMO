using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordDocument_SaveAsPdf_MetadataOverrides() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfMetadataOverride.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfMetadataOverride.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Properties.Title = "Original Title";
                document.Properties.Creator = "Original Author";
                document.Properties.Subject = "Original Subject";
                document.Properties.Keywords = "orig1, orig2";
                document.AddParagraph("Test");
                document.Save();
                var options = new WordPdfExportOptions {
                    Title = "Override Title",
                    Author = "Override Author",
                    Subject = "Override Subject",
                    Keywords = "override1, override2"
                };
                document.SaveAsPdf(pdfPath, options);
            }
            Assert.True(File.Exists(pdfPath));
            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                var info = pdf.Information;
                Assert.Equal("Override Title", info.Title);
                Assert.Equal("Override Author", info.Author);
                Assert.Equal("Override Subject", info.Subject);
                Assert.Equal("override1, override2", info.Keywords);
            }
        }
    }
}
