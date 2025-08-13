using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfAsync.docx");
        var pdfPath = Path.Combine(_directoryWithFiles, "PdfAsync.pdf");

        using (var document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            var saveTask = document.SaveAsPdfAsync(pdfPath, cancellationToken: CancellationToken.None);
            await saveTask;
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_ToStream() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfAsyncStream.docx");

        using (var document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            using (var stream = new MemoryStream()) {
                var saveTask = document.SaveAsPdfAsync(stream, cancellationToken: CancellationToken.None);
                await saveTask;
                Assert.True(stream.Length > 0);
            }
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public async Task Test_WordDocument_SaveAsPdfAsync_EmptyOrWhitespacePath_Throws(string path) {
        var docPath = Path.Combine(_directoryWithFiles, "PdfAsyncEmptyPath.docx");

        using (var document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();
            var ex = await Assert.ThrowsAsync<ArgumentException>(() => document.SaveAsPdfAsync(path, cancellationToken: CancellationToken.None));
            Assert.Contains("empty or whitespace", ex.Message);
        }
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_CanceledToken_Throws() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfAsyncCanceled.docx");
        var pdfPath = Path.Combine(_directoryWithFiles, "PdfAsyncCanceled.pdf");

        using (var document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();
            using (var cts = new CancellationTokenSource()) {
                cts.Cancel();
                await Assert.ThrowsAsync<OperationCanceledException>(() => document.SaveAsPdfAsync(pdfPath, cancellationToken: cts.Token));
            }
        }
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_ToStream_CanceledToken_Throws() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfAsyncStreamCanceled.docx");

        using (var document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();
            using (var stream = new MemoryStream()) {
                using (var cts = new CancellationTokenSource()) {
                    cts.Cancel();
                    await Assert.ThrowsAsync<OperationCanceledException>(() => document.SaveAsPdfAsync(stream, cancellationToken: cts.Token));
                }
            }
        }
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_ToByteArray() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfAsyncBytes.docx");

        using (var document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            byte[] bytes = await document.SaveAsPdfAsync(cancellationToken: CancellationToken.None);
            Assert.True(bytes.Length > 0);
        }
    }

    [Fact]
    public async Task Test_WordDocument_SaveAsPdfAsync_ToByteArray_CanceledToken_Throws() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfAsyncBytesCanceled.docx");

        using (var document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            using (var cts = new CancellationTokenSource()) {
                cts.Cancel();
                await Assert.ThrowsAsync<OperationCanceledException>(() => document.SaveAsPdfAsync(cancellationToken: cts.Token));
            }
        }
    }
}
