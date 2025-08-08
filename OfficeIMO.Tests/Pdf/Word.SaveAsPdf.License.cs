using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using QuestPDF.Infrastructure;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_DoesNotOverwriteExistingLicense() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfPreSetLicense.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfPreSetLicense.pdf");

            QuestPDF.Settings.License = LicenseType.Enterprise;
            try {
                using (WordDocument document = WordDocument.Create(docPath)) {
                    document.AddParagraph("Hello World");
                    document.Save();
                    document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                        QuestPdfLicenseType = LicenseType.Community
                    });
                }

                Assert.True(File.Exists(pdfPath));
                Assert.Equal(LicenseType.Enterprise, QuestPDF.Settings.License);
            } finally {
                QuestPDF.Settings.License = null;
            }
        }
    }
}

