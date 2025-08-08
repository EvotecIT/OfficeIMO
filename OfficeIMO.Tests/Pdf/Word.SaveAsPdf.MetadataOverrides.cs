using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordDocument_SaveAsPdf_MetadataOverrides() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfMetadataOverride.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfMetadataOverride.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.BuiltinDocumentProperties.Title = "Original Title";
                document.BuiltinDocumentProperties.Creator = "Original Author";
                document.BuiltinDocumentProperties.Subject = "Original Subject";
                document.BuiltinDocumentProperties.Keywords = "orig1, orig2";
                document.AddParagraph("Test");
                document.Save();
                var options = new PdfSaveOptions {
                    Title = "Override Title",
                    Author = "Override Author",
                    Subject = "Override Subject",
                    Keywords = "override1, override2"
                };
                document.SaveAsPdf(pdfPath, options);
            }
            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                var info = pdf.Information;
                Assert.Equal("Override Title", info.Title);
                Assert.Equal("Override Author", info.Author);
                Assert.Equal("Override Subject", info.Subject);
                Assert.Equal("override1, override2", info.Keywords);
            }
        }
    }
}
