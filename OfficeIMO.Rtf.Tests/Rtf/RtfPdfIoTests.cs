using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfIoTests {
    [Fact]
    public void RtfPdf_TypedDocument_WritesBytesStreamAndFile() {
        RtfDocument document = CreateDocument("Typed PDF");
        byte[] pdf = (document.ToPdf()).ToBytes();

        AssertPdfContains(pdf, "Typed PDF");

        using var output = new MemoryStream();
        output.WriteByte(0x2A);
        document.SaveAsPdf(output);
        Assert.True(output.CanWrite);
        AssertPdfContains(output.ToArray(), "Typed PDF");

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        try {
            document.SaveAsPdf(path);
            AssertPdfContains(File.ReadAllBytes(path), "Typed PDF");
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task RtfPdf_TypedDocument_AsyncWritesPerformIoAndHonorCancellation() {
        RtfDocument document = CreateDocument("Async PDF");
        using var output = new MemoryStream();

        await document.SaveAsPdfAsync(output);

        Assert.True(output.CanWrite);
        AssertPdfContains(output.ToArray(), "Async PDF");

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.SaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
    }

    [Fact]
    public async Task RtfPdf_SaveRejectsUnwritableStreamsAndKeepsCallerOwnership() {
        RtfDocument document = CreateDocument("Stream PDF");
        using var readOnly = new MemoryStream(Array.Empty<byte>(), writable: false);
        Assert.Throws<ArgumentException>(() => document.SaveAsPdf(readOnly));
        Assert.True(readOnly.CanRead);
        using var output = new MemoryStream();
        OfficeConversionReport report = await document.SaveAsPdfAsync(output);
        Assert.Equal("Pdf", report.Target);
        Assert.True(output.CanWrite);
        AssertPdfContains(output.ToArray(), "Stream PDF");
    }

    private static RtfDocument CreateDocument(string text) {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph(text);
        return document;
    }

    private static void AssertPdfContains(byte[] pdf, string expectedText) {
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(pdf, 0, 5));
        Assert.Contains(expectedText, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }
}
