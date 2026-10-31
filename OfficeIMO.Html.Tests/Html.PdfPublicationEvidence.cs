using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlPdfPublicationEvidenceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PublicPdfSavesRetainConversionEvidenceWhenTheDestinationRejectsBytes(bool asynchronous) {
        var source = HtmlConversionDocument.Parse("<p style='font-family:Helvetica'>Publication evidence</p>");
        using var output = new RejectingStream();
        OfficeConversionException error;
        if (asynchronous) error = await Assert.ThrowsAsync<OfficeConversionException>(() => source.SaveAsPdfAsync(output));
        else error = Assert.Throws<OfficeConversionException>(() => source.SaveAsPdf(output));
        var report = Assert.IsType<OfficeConversionReport>(error.Report);
        Assert.Equal("Html", report.Source);
        Assert.Equal("Pdf", report.Target);
        Assert.Contains(report.Diagnostics, d => d.Code == OfficeDiagnosticCodes.ConversionOutputFailed);
        Assert.IsType<IOException>(error.InnerException);
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
