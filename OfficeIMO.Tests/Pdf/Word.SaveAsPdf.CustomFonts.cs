using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordDocument_SaveAsPdf_CustomFontFile() {
            string fontPath = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf")
                : RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                    ? "/System/Library/Fonts/Supplemental/Arial.ttf"
                    : "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf";
            string expectedFont = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ||
                                   RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                ? "Arial"
                : "DejaVuSans";
            string docPath = Path.Combine(_directoryWithFiles, "PdfFontFile.docx");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello from file font").FontFamily = "FileFont";
                document.Save();
                var options = new PdfSaveOptions {
                    FontFilePaths = new Dictionary<string, string> { { "FileFont", fontPath } }
                };
                byte[] pdf = document.SaveAsPdf(options);
                using (PdfDocument pdfDoc = PdfDocument.Open(new MemoryStream(pdf))) {
                    var fonts = pdfDoc.GetPage(1).Letters.Select(l => l.FontName).Distinct();
                    Assert.Contains(fonts, f => f != null && f.Contains(expectedFont));
                }
            }
        }

        [Fact]
        public void Test_WordDocument_SaveAsPdf_CustomFontStream() {
            string fontPath = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf")
                : RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                    ? "/System/Library/Fonts/Supplemental/Arial.ttf"
                    : "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf";
            string expectedFont = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ||
                                   RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                ? "Arial"
                : "DejaVuSans";
            string docPath = Path.Combine(_directoryWithFiles, "PdfFontStream.docx");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello from stream font").FontFamily = "StreamFont";
                document.Save();
                using var fs = File.OpenRead(fontPath);
                var options = new PdfSaveOptions {
                    FontStreams = new Dictionary<string, Stream> { { "StreamFont", fs } }
                };
                byte[] pdf = document.SaveAsPdf(options);
                using (PdfDocument pdfDoc = PdfDocument.Open(new MemoryStream(pdf))) {
                    var fonts = pdfDoc.GetPage(1).Letters.Select(l => l.FontName).Distinct();
                    Assert.Contains(fonts, f => f != null && f.Contains(expectedFont));
                }
            }
        }

        [Fact]
        public void Test_WordDocument_SaveAsPdf_CustomFontFile_CanRetryAfterMissingPath() {
            string fontPath = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf")
                : RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                    ? "/System/Library/Fonts/Supplemental/Arial.ttf"
                    : "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf";
            string expectedFont = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ||
                                   RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                ? "Arial"
                : "DejaVuSans";
            string docPath = Path.Combine(_directoryWithFiles, "PdfFontRetry.docx");
            string fontAlias = "RetryFont" + Guid.NewGuid().ToString("N");
            string missingPath = Path.Combine(_directoryWithFiles, "missing-font-file.ttf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello from retried file font").FontFamily = fontAlias;
                document.Save();

                byte[] firstPdf = document.SaveAsPdf(new PdfSaveOptions {
                    FontFilePaths = new Dictionary<string, string> { { fontAlias, missingPath } }
                });
                Assert.NotEmpty(firstPdf);

                byte[] secondPdf = document.SaveAsPdf(new PdfSaveOptions {
                    FontFilePaths = new Dictionary<string, string> { { fontAlias, fontPath } }
                });

                using (PdfDocument pdfDoc = PdfDocument.Open(new MemoryStream(secondPdf))) {
                    var fonts = pdfDoc.GetPage(1).Letters.Select(l => l.FontName).Distinct();
                    Assert.Contains(fonts, f => f != null && f.Contains(expectedFont));
                }
            }
        }

        [Fact]
        public void Test_WordDocument_SaveAsPdf_CustomFontStream_RewindsSourceAfterFailure() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfFontStreamFailure.docx");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello from invalid stream font").FontFamily = "BrokenStreamFont";
                document.Save();

                using MemoryStream invalidFontStream = new MemoryStream(new byte[] { 1, 2, 3, 4 });
                Assert.ThrowsAny<Exception>(() => document.SaveAsPdf(new PdfSaveOptions {
                    FontStreams = new Dictionary<string, Stream> { { "BrokenStreamFont", invalidFontStream } }
                }));

                Assert.Equal(0, invalidFontStream.Position);
            }
        }
    }
}
