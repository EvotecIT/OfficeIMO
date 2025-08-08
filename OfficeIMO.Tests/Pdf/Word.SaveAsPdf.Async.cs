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

            await document.SaveAsPdfAsync(pdfPath);
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
                await document.SaveAsPdfAsync(stream);
                Assert.True(stream.Length > 0);
            }
        }
    }
    }

