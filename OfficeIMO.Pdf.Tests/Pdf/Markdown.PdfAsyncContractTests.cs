using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class MarkdownPdfAsyncContractTests {
    [Fact]
    public async Task PdfAsyncSavesPerformIoAndDoNotTurnCancellationIntoFailureResults() {
        MarkdownDoc document = MarkdownDoc.Create().H1("Async Markdown PDF");
        using var output = new MemoryStream();

        await document.SaveAsPdfAsync(output);

        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(output.ToArray(), 0, 5));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.SaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.TrySaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
    }
}
