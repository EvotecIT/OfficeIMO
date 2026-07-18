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
        byte[] pdf = document.ToPdf();

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
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.TrySaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
    }

    [Fact]
    public async Task RtfPdf_TrySave_ReturnsWriteFailuresButNotCancellation() {
        RtfDocument document = CreateDocument("Try PDF");
        using var readOnly = new MemoryStream(Array.Empty<byte>(), writable: false);

        PdfCore.PdfSaveResult failure = document.TrySaveAsPdf(readOnly);

        Assert.False(failure.Succeeded);
        Assert.NotEmpty(failure.Diagnostics);

        using var output = new MemoryStream();
        PdfCore.PdfSaveResult success = await document.TrySaveAsPdfAsync(output);
        Assert.True(success.Succeeded);
        Assert.Equal(output.Length, success.BytesWritten);
    }

    [Fact]
    public void RtfPdf_PublicAdapterSurface_RequiresParsedSourceModels() {
        MethodInfo[] methods = typeof(RtfPdfConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static);

        Assert.DoesNotContain(methods, method => method.Name.Contains("FromRtf", StringComparison.Ordinal));
        Assert.All(
            methods.Where(method => method.Name != nameof(RtfPdfConverterExtensions.ToRtfDocument)),
            method => Assert.Equal(typeof(RtfDocument), method.GetParameters()[0].ParameterType));

        MethodInfo import = Assert.Single(methods, method => method.Name == nameof(RtfPdfConverterExtensions.ToRtfDocument));
        Assert.Equal(typeof(PdfCore.PdfLogicalDocument), import.GetParameters()[0].ParameterType);
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
