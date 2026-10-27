using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfConversionPublicationEvidenceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LosslessPublicationFailureRetainsCompletedSourceEvidence(bool asynchronous) {
        using var document = PdfDocument.Create().Paragraph(p => p.Text("Publication evidence"));
        var source = new OfficeConversionReport("Word", "Pdf", diagnostics: new[] {
            new OfficeDiagnostic(OfficeDiagnosticCodes.WordToMarkdownWarning, OfficeDiagnosticSeverity.Information,
                "Source was inspected.", lossKind: OfficeConversionLossKind.None)
        });
        var conversion = new PdfDocumentConversionResult(document, new PdfConversionReport()).WithSourceConversionReport(source);
        using var output = new RejectingStream();
        OfficeConversionException error;
        if (asynchronous) error = await Assert.ThrowsAsync<OfficeConversionException>(() =>
            PdfConversionEvidence.SaveLosslessAsync(conversion, "Word", output, Stopwatch.StartNew(), default));
        else error = Assert.Throws<OfficeConversionException>(() =>
            PdfConversionEvidence.SaveLossless(conversion, "Word", output, Stopwatch.StartNew(), default));
        var report = Assert.IsType<OfficeConversionReport>(error.Report);
        Assert.Contains(report.Diagnostics, d => d.Message == "Source was inspected.");
        Assert.Contains(report.Diagnostics, d => d.Code == OfficeDiagnosticCodes.ConversionOutputFailed);
        Assert.IsType<IOException>(error.InnerException);
        Assert.Null(report.OutputPath);
        Assert.True(output.CanWrite);
    }

    private sealed class RejectingStream : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("Destination rejected the bytes.");
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Task.FromException(new IOException("Destination rejected the bytes."));
#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) => throw new IOException("Destination rejected the bytes.");
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) =>
            new(Task.FromException(new IOException("Destination rejected the bytes.")));
#endif
    }
}
