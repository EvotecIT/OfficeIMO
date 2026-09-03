using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlProvenanceWave77Tests {
    [Fact]
    public async Task DocumentParserRejectsCancellationAtTheWorkerBoundary() {
        const string html = "<!doctype html><html><body>content</body></html>";
        using var cancellation = new CancellationTokenSource();
        var enteredWorker = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseWorker = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        Task parse = Task.Run(async () => {
            enteredWorker.SetResult(true);
            await releaseWorker.Task;
            HtmlDocumentParser.ParseDocument(html, cancellation.Token);
        });
        await enteredWorker.Task;
        cancellation.Cancel();
        releaseWorker.SetResult(true);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => parse);
    }
}
