using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderNativeBatchInputBudgetTests {
    [Theory]
    [InlineData(".pdf", false)]
    [InlineData(".docx", false)]
    [InlineData(".doc", false)]
    [InlineData(".xlsx", false)]
    [InlineData(".xls", false)]
    [InlineData(".pptx", false)]
    [InlineData(".ppt", false)]
    [InlineData(".pdf", true)]
    [InlineData(".docx", true)]
    [InlineData(".xlsx", true)]
    [InlineData(".pptx", true)]
    public async Task NativeLoadEnforcesRemainingAllowanceWhenAPathGrowsAfterInitialAdmission(string extension, bool asynchronous) {
        byte[] bytes = Source(extension);
        Assert.True(bytes.Length > 256);
        string path = Path.Combine(Path.GetTempPath(), "reader-native-budget-" + Guid.NewGuid().ToString("N") + extension);
        try {
            File.WriteAllBytes(path, new byte[128]);
            bool entered = false;
            OfficeDocumentReadResult Read(string source, ReaderOptions options, CancellationToken token) {
                // This registered boundary runs after Core's path admission and before
                // the real format loader. Replacing a source here makes growth deterministic.
                entered = true;
                File.WriteAllBytes(source, bytes);
                return ReadNative(extension, source, options, token);
            }
            var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
                Id = "native-boundary", Extensions = new[] { extension },
                ReadDocumentPath = asynchronous ? null : Read,
                ReadDocumentPathAsync = asynchronous ? (source, options, token) => Task.FromResult(Read(source, options, token)) : null
            }).Build();
            Exception? error = await Record.ExceptionAsync(() => reader.ReadDocumentsAsync(new[] { path },
                batchOptions: new ReaderBatchOptions { MaxTotalBytes = 256 }));
            Assert.True(entered);
            Assert.NotNull(error);
            if (error is OfficePackageSecurityException package) {
                Assert.Equal(OfficePackageSecurityRule.PackageSize, package.Rule);
                Assert.Equal(256d, package.Limit);
            } else if (extension == ".pdf") {
                var limit = Assert.IsType<ReaderAdmissionLimitException>(error);
                Assert.Equal("Pdf.InputBytes", limit.LimitName);
                Assert.Equal(256d, limit.Maximum);
                Assert.IsType<PdfReadLimitException>(limit.InnerException);
            } else {
                var limit = Assert.IsType<ReaderResourceLimitException>(error);
                Assert.EndsWith("MaxInputBytes", limit.LimitName, StringComparison.Ordinal);
                Assert.Equal(256, limit.Maximum);
            }
        } finally { File.Delete(path); }
    }

    private static OfficeDocumentReadResult ReadNative(string extension, string path, ReaderOptions options, CancellationToken token) {
        if (extension == ".pdf") {
            using var source = File.OpenRead(path);
            return PdfReaderAdapter.ReadDocument(source, path, options, null, token);
        }
        return extension switch {
        ".docx" or ".doc" => WordReaderAdapter.ReadDocument(path, options, new ReaderWordOptions(), null, false, token),
        ".xlsx" or ".xls" => ExcelReaderAdapter.ReadDocument(path, options, new ReaderExcelOptions(), token),
        _ => PowerPointReaderAdapter.ReadDocument(path, options, new ReaderPowerPointOptions(), token)
        };
    }

    private static byte[] Source(string extension) {
        if (extension == ".pdf") return PdfDocument.Create().H1("Bounded native input").ToBytes();
        if (extension == ".docx" || extension == ".doc") {
            using var document = WordDocument.Create(); document.AddParagraph("Bounded native input");
            return document.ToBytes(new WordSaveOptions { Format = extension == ".doc" ? WordFileFormat.Doc : WordFileFormat.Docx });
        }
        if (extension == ".xlsx" || extension == ".xls") {
            using var document = ExcelDocument.Create(); document.AddWorksheet("Evidence").CellValue(1, 1, "Bounded native input");
            return document.ToBytes(new ExcelSaveOptions { Format = extension == ".xls" ? ExcelFileFormat.Xls : ExcelFileFormat.Xlsx });
        }
        using var presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddTextBox("Bounded native input", 100000, 100000, 2000000, 500000);
        return presentation.ToBytes(new PowerPointSaveOptions { Format = extension == ".ppt" ? PowerPointFileFormat.Ppt : PowerPointFileFormat.Pptx });
    }
}
