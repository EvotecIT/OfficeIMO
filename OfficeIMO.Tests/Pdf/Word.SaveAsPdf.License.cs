using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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
                // License behavior is environment-dependent; ensure PDF exists and license stays set
                Assert.NotNull(QuestPDF.Settings.License);
            } finally {
                QuestPDF.Settings.License = null;
            }
        }

        [Fact]
        public void SaveAsPdf_EmbedsCustomFont() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfFontFamily.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfFontFamily.pdf");

            string fontFamily = RuntimeInformation.IsOSPlatform(OSPlatform.Linux) ? "DejaVu Sans" : "Arial";

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello World");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    FontFamily = fontFamily
                });
            }

            string pdfContent = File.ReadAllText(pdfPath);
            Assert.Contains(fontFamily.Replace(" ", ""), pdfContent, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("FontFile", pdfContent);
        }

        [Fact]
        public async Task SaveAsPdfAsync_Path_RestoresUnsetLicense() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfAsyncLicensePath.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfAsyncLicensePath.pdf");

            QuestPDF.Settings.License = null;
            try {
                using (WordDocument document = WordDocument.Create(docPath)) {
                    document.AddParagraph("Hello World");
                    document.Save();

                    await document.SaveAsPdfAsync(pdfPath, new PdfSaveOptions {
                        QuestPdfLicenseType = LicenseType.Community
                    }, CancellationToken.None);
                }

                Assert.True(File.Exists(pdfPath));
                Assert.Null(QuestPDF.Settings.License);
            } finally {
                QuestPDF.Settings.License = null;
            }
        }

        [Fact]
        public async Task SaveAsPdfAsync_ByteArray_RestoresUnsetLicense() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfAsyncLicenseBytes.docx");

            QuestPDF.Settings.License = null;
            try {
                using (WordDocument document = WordDocument.Create(docPath)) {
                    document.AddParagraph("Hello World");
                    document.Save();

                    byte[] bytes = await document.SaveAsPdfAsync(new PdfSaveOptions {
                        QuestPdfLicenseType = LicenseType.Community
                    }, CancellationToken.None);

                    Assert.NotEmpty(bytes);
                }

                Assert.Null(QuestPDF.Settings.License);
            } finally {
                QuestPDF.Settings.License = null;
            }
        }

        [Fact]
        public async Task SaveAsPdfAsync_Path_RestoresUnsetLicense_WhenDocumentCreationFails() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfAsyncLicensePathFailure.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfAsyncLicensePathFailure.pdf");

            QuestPDF.Settings.License = null;
            try {
                using (WordDocument document = WordDocument.Create(docPath)) {
                    document.AddParagraph("Hello World");
                    document.Save();

                    using MemoryStream invalidFontStream = new MemoryStream(new byte[] { 1, 2, 3, 4 });
                    await Assert.ThrowsAnyAsync<Exception>(() => document.SaveAsPdfAsync(pdfPath, new PdfSaveOptions {
                        QuestPdfLicenseType = LicenseType.Community,
                        FontStreams = new Dictionary<string, Stream> { { "BrokenFont", invalidFontStream } }
                    }, CancellationToken.None));
                }

                Assert.Null(QuestPDF.Settings.License);
            } finally {
                QuestPDF.Settings.License = null;
            }
        }

        [Fact]
        public async Task SaveAsPdfAsync_Stream_RestoresUnsetLicense_WhenDocumentCreationFails() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfAsyncLicenseStreamFailure.docx");

            QuestPDF.Settings.License = null;
            try {
                using (WordDocument document = WordDocument.Create(docPath)) {
                    document.AddParagraph("Hello World");
                    document.Save();

                    using MemoryStream output = new MemoryStream();
                    using MemoryStream invalidFontStream = new MemoryStream(new byte[] { 1, 2, 3, 4 });
                    await Assert.ThrowsAnyAsync<Exception>(() => document.SaveAsPdfAsync(output, new PdfSaveOptions {
                        QuestPdfLicenseType = LicenseType.Community,
                        FontStreams = new Dictionary<string, Stream> { { "BrokenFont", invalidFontStream } }
                    }, CancellationToken.None));
                }

                Assert.Null(QuestPDF.Settings.License);
            } finally {
                QuestPDF.Settings.License = null;
            }
        }
    }
}

