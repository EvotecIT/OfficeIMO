using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlConversionOptionsCaptureTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ResourceAwaitKeepsTheOriginalLossPolicyAndProtectsTheDestination(bool save) {
        var entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<HtmlResolvedResource?>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        var options = new HtmlToPdfOptions {
            FailOnLoss = true,
            ResourcePolicy = new OfficeIMO.Pdf.PdfResourcePolicy { AllowRemoteResourceResolution = true },
            ResourceResolver = (_, _) => {
                entered.TrySetResult(true);
                return release.Task;
            }
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://example.test/missing.css'><p>Retained body</p>");
        using var destination = new MemoryStream();
        destination.Write(new byte[] { 1, 2, 3 });
        Task operation = save ? source.SaveAsPdfAsync(destination, options, cancellation.Token) : ConvertAndDispose();
        try {
            await entered.Task.WaitAsync(cancellation.Token);
            options.FailOnLoss = false;
            release.SetResult(null);
            await Assert.ThrowsAsync<OfficeConversionException>(() => operation);
            Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
            Assert.True(destination.CanWrite);
        } finally { release.TrySetResult(null); }

        async Task ConvertAndDispose() {
            var converted = await source.ToPdfAsync(options, cancellation.Token);
            converted.Value.Dispose();
        }
    }
}
