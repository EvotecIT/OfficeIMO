using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
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
            Assert.False(saveTask.IsCompleted, "SaveAsPdfAsync should not complete synchronously");
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
                Assert.False(saveTask.IsCompleted, "SaveAsPdfAsync should not complete synchronously");
                await saveTask;
                Assert.True(stream.Length > 0);
            }
        }
    }
    }

