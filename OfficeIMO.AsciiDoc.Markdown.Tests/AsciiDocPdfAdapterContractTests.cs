using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.AsciiDoc.Markdown.Tests;

public sealed class AsciiDocPdfAdapterContractTests {
    [Fact]
    public void ParserDiagnostics_FlowIntoFinalPdfResult() {
        AsciiDocDocument document = AsciiDocDocument.Parse("= Parser proof\n\n----\nunterminated block\n").Document;

        var result = document.ToPdfDocumentResult();
        var warning = Assert.Single(result.Warnings, item => item.Code == "ADOC001");

        Assert.Equal("OfficeIMO.AsciiDoc.Pdf", warning.Converter);
        Assert.Equal(OfficeIMO.Pdf.PdfConversionWarningSeverity.Error, warning.Severity);
        Assert.Equal("parse", warning.Details["stage"]);
        Assert.True(result.HasLoss);
    }

    [Fact]
    public void SaveAsPdf_LeavesCallerOwnedStreamOpen() {
        AsciiDocDocument document = AsciiDocDocument.Parse("= Stream proof\n\nCaller ownership marker.\n").Document;
        using var stream = new MemoryStream();

        document.SaveAsPdf(stream);
        long written = stream.Length;
        stream.Position = stream.Length;
        stream.WriteByte(0);

        Assert.True(written > 100);
        Assert.Equal(written + 1, stream.Length);
    }

    [Fact]
    public async Task SaveAsPdfAsync_PreCanceled_DoesNotConvertOrWrite() {
        AsciiDocDocument document = AsciiDocDocument.Parse("= Cancellation proof\n").Document;
        using var stream = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.SaveAsPdfAsync(stream, cancellationToken: cancellation.Token));

        Assert.Equal(0, stream.Length);
        Assert.True(stream.CanWrite);
    }
}
