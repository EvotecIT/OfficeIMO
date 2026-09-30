using System.Collections;
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
