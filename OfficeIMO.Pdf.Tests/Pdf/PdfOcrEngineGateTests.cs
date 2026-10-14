using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

// The provider-release handshake needs thread-pool continuations. Run it apart
// from rendering collections that can occupy those workers on .NET Framework.
[Collection(PdfOcrEngineGateCollection.Name)]
public sealed class PdfOcrEngineGateTests {
    [Fact]
    public async Task RecognizeAndMergeAsync_DoesNotRestartTimedOutNonConcurrentProvider() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 220, 90)
            .ToBytes();
        var completion = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var firstCallStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        int callCount = 0;
        var provider = new DelegateOcrEngine("stalled-non-concurrent-pdf-fixture", (_, _) => {
            Interlocked.Increment(ref callCount);
            firstCallStarted.TrySetResult(null);
            return completion.Task;
        });
        // Isolate the provider-start/release handshake from parallel rendering
        // collections while keeping the same real provider and gate deadlines.
        var options = new PdfOcrMergeOptions { ProviderTimeout = TimeSpan.FromSeconds(5) };
        try {
            Task<PdfOcrMergeResult> firstCall = PdfDocument.Load(pdf).ReadWithOcrAsync(provider, options);
            Task started = await Task.WhenAny(firstCallStarted.Task, firstCall);
            if (started != firstCallStarted.Task) await firstCall; // Preserve the actual failure when startup fails.
            Assert.Same(firstCallStarted.Task, started);
            OcrEngineTimeoutException first = await Assert.ThrowsAsync<OcrEngineTimeoutException>(() => firstCall);
            var occupiedGateOptions = new PdfOcrMergeOptions { ProviderTimeout = TimeSpan.FromMilliseconds(200) };
            OcrEngineTimeoutException second = await Assert.ThrowsAsync<OcrEngineTimeoutException>(
                () => PdfDocument.Load(pdf).ReadWithOcrAsync(provider, occupiedGateOptions));
            Assert.True(first.ProviderCallStarted);
            Assert.False(second.ProviderCallStarted);
            Assert.Equal(1, Volatile.Read(ref callCount));
        } finally { completion.TrySetResult(new OcrResult()); }
    }

    [Fact]
    public async Task OrientationAndRecognitionShareTheNonConcurrentEngineGate() {
        var firstEntered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        int calls = 0;
        var engine = new DelegateOcrEngine("shared-gate", async (request, _) => {
            int call = Interlocked.Increment(ref calls);
            if (call == 1) { firstEntered.SetResult(true); await release.Task.ConfigureAwait(false); }
            return new OcrResult();
        }, new OcrEngineCapabilities { SupportsOrientationDetection = true, SupportsConcurrentRequests = false });
        Task<OcrResult> orientation = OcrEngineRunner.RecognizeAsync(engine,
            new OcrRequest { Operation = OcrOperation.DetectOrientation }, TimeSpan.FromSeconds(10));
#pragma warning disable xUnit1030 // The provider handshake deliberately bypasses the test synchronization context.
        await firstEntered.Task.ConfigureAwait(false);
        Task<OcrResult> recognition;
        try {
            recognition = OcrEngineRunner.RecognizeAsync(engine, new OcrRequest(), TimeSpan.FromSeconds(10));
            Assert.Equal(1, Volatile.Read(ref calls));
        } finally { release.TrySetResult(true); }
        await Task.WhenAll(orientation, recognition).ConfigureAwait(false);
#pragma warning restore xUnit1030
        Assert.Equal(2, calls);
    }
}

[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class PdfOcrEngineGateCollection {
    public const string Name = "PDF OCR shared engine gate";
}
