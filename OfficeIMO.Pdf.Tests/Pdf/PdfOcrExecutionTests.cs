using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOcrExecutionTests {
    [Theory]
    [InlineData(true, 2)]
    [InlineData(false, 1)]
    public async Task Ocr_BoundsOutstandingPagesAndRetainsSelectionOrder(bool concurrent, int expectedOutstanding) {
        var engine = new ControlledEngine(concurrent);
        var document = PdfDocument.Create().Paragraph(p => p.Text("One")).PageBreak()
            .Paragraph(p => p.Text("Two")).PageBreak().Paragraph(p => p.Text("Three"));
        Task<PdfOcrMergeResult> operation = document.ReadWithOcrAsync(engine, new PdfOcrMergeOptions {
            Dpi = 36,
            MaxConcurrentPages = 2,
            ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(3, 1, 2) }
        });
        await engine.WaitForCallsAsync(expectedOutstanding);
        Assert.Equal(expectedOutstanding, engine.CallCount);
        if (concurrent) {
            engine.Complete(1);
            await engine.WaitForCallsAsync(3);
            Assert.Equal(2, engine.ActiveCalls);
            engine.Complete(2);
            engine.Complete(3);
        } else {
            engine.Complete(3);
            await engine.WaitForCallsAsync(2);
            engine.Complete(1);
            await engine.WaitForCallsAsync(3);
            engine.Complete(2);
        }
        PdfOcrMergeResult result = await operation;
        Assert.Equal(new[] { 3, 1, 2 }, result.Pages.Select(page => page.PageNumber));
        Assert.Equal(0, engine.ActiveCalls);
    }

    [Fact]
    public async Task Ocr_CancellationReachesOutstandingProviders() {
        var engine = new ControlledEngine(true);
        var document = PdfDocument.Create().Paragraph(p => p.Text("One")).PageBreak().Paragraph(p => p.Text("Two"));
        using var cancellation = new CancellationTokenSource();
        Task<PdfOcrMergeResult> operation = document.ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { Dpi = 36, MaxConcurrentPages = 2 }, cancellationToken: cancellation.Token);
        await engine.WaitForCallsAsync(2);
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
        await engine.WaitForIdleAsync();
        Assert.Equal(0, engine.ActiveCalls);
    }

    [Fact]
    public async Task Ocr_PreservesSynchronousProviderFailureWhenCancellingRemainingWork() {
        var engine = new DelegateOcrEngine("failure", (request, token) =>
            Task.FromException<OcrResult>(new InvalidOperationException("Provider unavailable")),
            new OcrEngineCapabilities { SupportsConcurrentRequests = true, SupportedMediaTypes = new[] { "image/png" } });
        var document = PdfDocument.Create().Paragraph(p => p.Text("One")).PageBreak().Paragraph(p => p.Text("Two"));
        var failure = await Assert.ThrowsAsync<InvalidOperationException>(() => document.ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { Dpi = 36, MaxConcurrentPages = 2 }));
        Assert.Contains("Provider unavailable", failure.Message);
    }

    [Fact]
    public async Task Ocr_PreservesRenderWarningsWithProviderDiagnostics() {
        var document = PdfDocument.Create().Paragraph(p => p.Text("Font fallback"));
        string[] renderWarnings = document.Render.Pages().Single().Diagnostics.ToArray();
        Assert.NotEmpty(renderWarnings);
        var engine = new ControlledEngine(false, completeImmediately: true);
        PdfOcrMergeResult result = await document.ReadWithOcrAsync(engine);
        PdfOcrPageMergeResult page = Assert.Single(result.Pages);
        foreach (string warning in renderWarnings) Assert.Contains(warning, page.Diagnostics);
        Assert.Contains("provider: fixture", page.Diagnostics);
    }

    [Fact]
    public async Task Ocr_RejectsOversizedPngBeforeCallingProvider() {
        var engine = new ControlledEngine(false, completeImmediately: true);
        var document = PdfDocument.Create().Paragraph(p => p.Text("Page"));
        await Assert.ThrowsAsync<PdfReadLimitException>(() => document.ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { MaxRenderedBytesPerPage = 8 }));
        Assert.Equal(0, engine.CallCount);
    }

    private sealed class ControlledEngine : IOcrEngine {
        private readonly object _gate = new object();
        private readonly Dictionary<int, TaskCompletionSource<OcrResult>> _requests = new();
        private readonly bool _completeImmediately;
        private int _calls;
        private int _active;
        private readonly TaskCompletionSource<bool> _idle = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal ControlledEngine(bool concurrent, bool completeImmediately = false) {
            _completeImmediately = completeImmediately;
            Capabilities = new OcrEngineCapabilities {
                SupportsConcurrentRequests = concurrent, SupportedMediaTypes = new[] { "image/png" }
            };
        }
        public string Id => "controlled";
        public OcrEngineCapabilities Capabilities { get; }
        internal int CallCount => Volatile.Read(ref _calls);
        internal int ActiveCalls => Volatile.Read(ref _active);
        public async Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
            var completion = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (_gate) _requests.Add(request.PageNumber!.Value, completion);
            Interlocked.Increment(ref _active);
            Interlocked.Increment(ref _calls);
            using var registration = cancellationToken.Register(() => completion.TrySetCanceled());
            try {
                if (_completeImmediately) Complete(request.PageNumber.Value);
                return await completion.Task.ConfigureAwait(false);
            } finally {
                if (Interlocked.Decrement(ref _active) == 0) _idle.TrySetResult(true);
            }
        }
        internal void Complete(int page) {
            lock (_gate) _requests[page].TrySetResult(new OcrResult {
                Diagnostics = new[] { new OcrDiagnostic { Code = "provider", Message = "fixture" } }
            });
        }
        internal async Task WaitForCallsAsync(int count) {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            while (CallCount < count) await Task.Delay(10, timeout.Token);
        }
        internal async Task WaitForIdleAsync() {
            // The shared runner returns promptly on cancellation even when a provider ignores it.
            // A cooperative provider settles asynchronously; require that settlement without a race.
            Assert.Same(_idle.Task, await Task.WhenAny(_idle.Task, Task.Delay(TimeSpan.FromSeconds(30))));
            await _idle.Task;
        }
    }
}
