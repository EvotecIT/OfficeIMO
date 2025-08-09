using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
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

            var saveTask = document.SaveAsPdfAsync(pdfPath);
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
                var saveTask = document.SaveAsPdfAsync(stream);
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
            var ex = await Assert.ThrowsAsync<ArgumentException>(() => document.SaveAsPdfAsync(path));
            Assert.Contains("empty or whitespace", ex.Message);
        }
    }
}