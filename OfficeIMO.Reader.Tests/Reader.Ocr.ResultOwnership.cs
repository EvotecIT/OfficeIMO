using System.Collections;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrResultOwnershipTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task ProviderCollectionFailuresAreContentFree(int collection) {
        var result = new OcrResult();
        if (collection == 1) result.Spans = new ThrowingList<OcrTextSpan>();
        else if (collection == 2) result.Diagnostics = new[] { new OcrDiagnostic { Attributes = new ThrowingAttributes() } };
        else result.Diagnostics = new ThrowingList<OcrDiagnostic>();
        var engine = new DelegateOcrEngine("collection", (_, _) => Task.FromResult(result));
        var error = await Assert.ThrowsAsync<OcrEngineExecutionException>(() =>
            OcrEngineRunner.RecognizeAsync(engine, new OcrRequest(), TimeSpan.FromSeconds(5)));
        Assert.Equal(OcrEngineFailureKind.ProviderFailure, error.Kind);
        Assert.Null(error.InnerException);
        Assert.DoesNotContain("FAKE_AUDIT_SENTINEL", error.ToString());
    }

    [Fact]
    public async Task CapturedResultDoesNotRetainProviderOwnedCollectionsOrGeometry() {
        var attributes = new Dictionary<string, string> { ["page"] = "1" };
        var span = new OcrTextSpan { Text = "original", Region = new OcrRegion { X = 12 } };
        var result = new OcrResult {
            Spans = new[] { span }, Diagnostics = new[] { new OcrDiagnostic { Attributes = attributes } }
        };
        var engine = new DelegateOcrEngine("ownership", (_, _) => Task.FromResult(result));
        OcrResult captured = await OcrEngineRunner.RecognizeAsync(engine, new OcrRequest(), TimeSpan.FromSeconds(5));
        span.Text = "changed"; span.Region.X = 99; attributes["page"] = "changed";
        Assert.Equal("original", Assert.Single(captured.Spans).Text);
        Assert.Equal(12, captured.Spans[0].Region!.X);
        Assert.Equal("1", Assert.Single(captured.Diagnostics).Attributes["page"]);
    }

    [Fact]
    public async Task CaptureReadsOnlyBoundedSpanAndAttributePrefixes() {
        var spans = new PrefixList<OcrTextSpan>(1_000_000,
            new[] { new OcrTextSpan { Text = "first" }, new OcrTextSpan { Text = "second" } });
        var attributes = new PrefixAttributes();
        var engine = new DelegateOcrEngine("bounded", (_, _) => Task.FromResult(new OcrResult {
            Spans = spans, Diagnostics = new[] { new OcrDiagnostic { Attributes = attributes } }
        }));
        OcrResult captured = await OcrEngineRunner.CreateExecution(engine).RecognizeAsync(
            new OcrRequest(), TimeSpan.FromSeconds(5), new OcrResultCaptureLimits(2, 1, 2), CancellationToken.None);
        Assert.Equal(2, captured.Spans.Count);
        Assert.Equal(999_998, captured.OmittedSpanCount);
        Assert.Equal(2, captured.Diagnostics[0].Attributes.Count);
        Assert.Equal(999_998, captured.Diagnostics[0].OmittedAttributeCount);
    }

    [Fact]
    public async Task TerminalDiagnosticOutsideRetainedPrefixStillRejectsRecognition() {
        var engine = new DelegateOcrEngine("terminal", (_, _) => Task.FromResult(new OcrResult {
            Diagnostics = new[] { new OcrDiagnostic(), new OcrDiagnostic { Severity = OcrDiagnosticSeverity.Error, IsRecoverable = false } }
        }));
        var error = await Assert.ThrowsAsync<OcrEngineExecutionException>(() =>
            OcrEngineRunner.CreateExecution(engine).RecognizeAsync(new OcrRequest(), TimeSpan.FromSeconds(5), new OcrResultCaptureLimits(1, 1, 1), CancellationToken.None));
        Assert.Equal(OcrEngineFailureKind.NonRecoverableDiagnostic, error.Kind);
    }

    [Fact]
    public async Task CancellationStopsRunnerOwnedCopyAndReleasesSharedGate() {
        using var cancellation = new CancellationTokenSource();
        int reads = 0;
        var spans = new CallbackList<OcrTextSpan>(1_000_000, () => {
            Interlocked.Increment(ref reads); cancellation.Cancel(); return new OcrTextSpan();
        });
        int calls = 0;
        var engine = new DelegateOcrEngine("canceled-copy", (_, _) => Task.FromResult(
            Interlocked.Increment(ref calls) == 1 ? new OcrResult { Spans = spans } : new OcrResult { Text = "next" }));
        var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            OcrEngineRunner.RecognizeAsync(engine, new OcrRequest(), TimeSpan.FromSeconds(5), cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        OcrResult next = await OcrEngineRunner.RecognizeAsync(engine, new OcrRequest(), TimeSpan.FromSeconds(5));
        Assert.Equal("next", next.Text);
        Assert.Equal(1, reads);
    }

    [Fact]
    public async Task ProviderCancellationDetailsRemainContentFree() {
        using var cancellation = new CancellationTokenSource();
        var engine = new DelegateOcrEngine("canceled-provider", (_, _) => {
            cancellation.Cancel();
            throw new OperationCanceledException("FAKE_AUDIT_SENTINEL", new Exception("FAKE_AUDIT_SENTINEL"), cancellation.Token);
        });
        var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            OcrEngineRunner.RecognizeAsync(engine, new OcrRequest(), TimeSpan.FromSeconds(5), cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Null(error.InnerException);
        Assert.DoesNotContain("FAKE_AUDIT_SENTINEL", error.ToString());
    }

    private sealed class PrefixList<T>(int count, T[] prefix) : IReadOnlyList<T> {
        public int Count => count;
        public T this[int index] => index < prefix.Length ? prefix[index] : throw new InvalidOperationException("Discarded span was inspected.");
        public IEnumerator<T> GetEnumerator() => throw new InvalidOperationException("Unbounded enumeration was attempted.");
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class CallbackList<T>(int count, Func<T> read) : IReadOnlyList<T> {
        public int Count => count;
        public T this[int index] => read();
        public IEnumerator<T> GetEnumerator() => throw new InvalidOperationException("Unbounded enumeration was attempted.");
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class PrefixAttributes : IReadOnlyDictionary<string, string> {
        public int Count => 1_000_000;
        public IEnumerable<string> Keys => throw new NotSupportedException();
        public IEnumerable<string> Values => throw new NotSupportedException();
        public string this[string key] => throw new NotSupportedException();
        public bool ContainsKey(string key) => throw new NotSupportedException();
        public bool TryGetValue(string key, out string value) => throw new NotSupportedException();
        public IEnumerator<KeyValuePair<string, string>> GetEnumerator() {
            yield return new("one", "1"); yield return new("two", "2");
            throw new InvalidOperationException("Discarded attributes were inspected.");
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class ThrowingAttributes : IReadOnlyDictionary<string, string> {
        public int Count => 1;
        public IEnumerable<string> Keys => Array.Empty<string>();
        public IEnumerable<string> Values => Array.Empty<string>();
        public string this[string key] => throw new InvalidOperationException("FAKE_AUDIT_SENTINEL");
        public bool ContainsKey(string key) => false;
        public bool TryGetValue(string key, out string value) { value = ""; return false; }
        public IEnumerator<KeyValuePair<string, string>> GetEnumerator() =>
            throw new InvalidOperationException("FAKE_AUDIT_SENTINEL");
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class ThrowingList<T> : IReadOnlyList<T> {
        public int Count => 1;
        public T this[int index] => throw new InvalidOperationException("FAKE_AUDIT_SENTINEL");
        public IEnumerator<T> GetEnumerator() => throw new InvalidOperationException("FAKE_AUDIT_SENTINEL");
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
