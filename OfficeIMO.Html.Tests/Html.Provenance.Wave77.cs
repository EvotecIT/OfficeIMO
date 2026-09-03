using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlProvenanceWave77Tests {
    [Fact]
    public async Task DocumentParserObservesCancellationDuringLargeParse() {
        string html = "<!doctype html><html><body>" +
                      string.Concat(Enumerable.Repeat("<div>content</div>", 100_000)) +
                      "</body></html>";
        using var cancellation = new CancellationTokenSource();

        Task parse = Task.Run(() => HtmlDocumentParser.ParseDocument(html, cancellation.Token));
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(10));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => parse);
    }
}
