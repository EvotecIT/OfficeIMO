using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.Diagnostics;
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

            var stopwatch = Stopwatch.StartNew();
            var saveTask = document.SaveAsPdfAsync(pdfPath);
            stopwatch.Stop();
            Assert.True(stopwatch.ElapsedMilliseconds < 100);
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
                var stopwatch = Stopwatch.StartNew();
                var saveTask = document.SaveAsPdfAsync(stream);
                stopwatch.Stop();
                Assert.True(stopwatch.ElapsedMilliseconds < 100);
                await saveTask;
                Assert.True(stream.Length > 0);
            }
        }
    }
    }

