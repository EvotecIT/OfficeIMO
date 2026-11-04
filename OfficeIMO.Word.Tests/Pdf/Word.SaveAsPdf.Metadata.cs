using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordDocument_SaveAsPdf_Metadata() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfDocumentProperties.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfDocumentProperties.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Properties.Title = "Pdf Title";
                document.Properties.Creator = "Pdf Author";
                document.Properties.Subject = "Pdf Subject";
                document.Properties.Keywords = "keyword1, keyword2";
                document.AddParagraph("Test");
                document.Save();
                document.SaveAsPdf(pdfPath);
            }

            Assert.True(File.Exists(pdfPath));

            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
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
                document.Properties.Title = "Native Pdf Title";
                document.Properties.Creator = "Native Pdf Author";
                document.Properties.Subject = "Native Pdf Subject";
                document.Properties.Keywords = "native, keyword";
                document.AddParagraph("Native metadata");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions());
            }

            Assert.True(File.Exists(pdfPath));

            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                var info = pdf.Information;
                Assert.Equal("Native Pdf Title", info.Title);
                Assert.Equal("Native Pdf Author", info.Author);
                Assert.Equal("Native Pdf Subject", info.Subject);
                Assert.Equal("native, keyword", info.Keywords);
            }
        }

        [Fact]
        public void Test_WordDocument_SaveAsPdf_OfficeIMOEngine_FontFamily_Does_Not_Pollute_MetadataKeywords() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeFontMetadata.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeFontMetadata.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Properties.Keywords = "native, keyword";
                document.AddParagraph("Native metadata font");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfExportOptions { FontFamily = "Times New Roman" });
            }

            using (PdfPigDocument pdf = PdfPigDocument.Open(pdfPath)) {
                Assert.Equal("native, keyword", pdf.Information.Keywords);
            }
        }
    }
}
