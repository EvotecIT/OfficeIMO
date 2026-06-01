using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using UglyToad.PdfPig;
using Xunit;
namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordDocument_SaveAsPdf_Metadata() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfMetadata.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfMetadata.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.BuiltinDocumentProperties.Title = "Pdf Title";
                document.BuiltinDocumentProperties.Creator = "Pdf Author";
                document.BuiltinDocumentProperties.Subject = "Pdf Subject";
                document.BuiltinDocumentProperties.Keywords = "keyword1, keyword2";
                document.AddParagraph("Test");
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                var info = pdf.Information;
                Assert.Equal("Pdf Title", info.Title);
                Assert.Equal("Pdf Author", info.Author);
                Assert.Equal("Pdf Subject", info.Subject);
                Assert.Equal("keyword1, keyword2", info.Keywords);
            }
        }

        [Fact]
        public void Test_WordDocument_SaveAsPdf_OfficeIMOEngine_Metadata() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeMetadata.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeMetadata.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.BuiltinDocumentProperties.Title = "Native Pdf Title";
                document.BuiltinDocumentProperties.Creator = "Native Pdf Author";
                document.BuiltinDocumentProperties.Subject = "Native Pdf Subject";
                document.BuiltinDocumentProperties.Keywords = "native, keyword";
                document.AddParagraph("Native metadata");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions());
            }

            Assert.True(File.Exists(pdfPath));

            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                var info = pdf.Information;
                Assert.Equal("Native Pdf Title", info.Title);
                Assert.Equal("Native Pdf Author", info.Author);
                Assert.Equal("Native Pdf Subject", info.Subject);
                Assert.Equal("native, keyword", info.Keywords);
            }
        }
    }
}
