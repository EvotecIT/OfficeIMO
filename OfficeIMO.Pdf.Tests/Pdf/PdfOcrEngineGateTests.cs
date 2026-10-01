using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

// The provider-release handshake needs thread-pool continuations. Run it apart
// from rendering collections that can occupy those workers on .NET Framework.
[Collection(PdfOcrEngineGateCollection.Name)]
public sealed class PdfOcrEngineGateTests {
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
