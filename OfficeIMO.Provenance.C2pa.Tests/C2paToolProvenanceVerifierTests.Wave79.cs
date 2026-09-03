using System.IO;
using System.Threading;

namespace OfficeIMO.Provenance.C2pa.Tests;

public sealed class C2paToolProvenanceVerifierTestsWave79 {
    [Fact]
    public void InterpretationBudgetReturnsAtTheDeadlineWhenWorkDoesNotCooperate() {
        using var release = new ManualResetEventSlim();
        var budget = new C2paToolExecutionBudget(TimeSpan.FromMilliseconds(20), CancellationToken.None);
        try {
            Assert.Throws<TimeoutException>(() => budget.RunInterpretation<int>(_ => {
                release.Wait();
                return 0;
            }));
        } finally {
            release.Set();
        }
    }

    [Fact]
    public void CancellationRemainsActiveAfterTheProviderProcessReturns() {
        string asset = CreateAsset();
        using var cancellation = new CancellationTokenSource();
        try {
            var verifier = new C2paToolProvenanceVerifier(
                "c2patool",
                new CallbackRunner(() => cancellation.Cancel()));

            Assert.Throws<OperationCanceledException>(() => verifier.Verify(
                asset,
                new OfficeProvenanceVerificationOptions(),
                cancellation.Token));
        } finally {
            File.Delete(asset);
        }
    }

    [Fact]
    public void ProviderTimeoutIncludesReportInterpretation() {
        string asset = CreateAsset();
        try {
            var verifier = new C2paToolProvenanceVerifier(
                "c2patool",
                new CallbackRunner(() => Thread.Sleep(30)));

            OfficeProvenanceVerificationResult result = verifier.Verify(
                asset,
                new OfficeProvenanceVerificationOptions { Timeout = TimeSpan.FromMilliseconds(10) });

            Assert.Equal(OfficeProvenanceVerificationStatus.Error, result.Status);
            Assert.Contains(result.Findings, finding => finding.Contains("interpretation", StringComparison.OrdinalIgnoreCase));
        } finally {
            File.Delete(asset);
        }
    }

    private static string CreateAsset() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".jpg");
        File.WriteAllBytes(path, new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 });
        return path;
    }

    private sealed class CallbackRunner : IC2paToolProcessRunner {
        private readonly Action _callback;

        internal CallbackRunner(Action callback) => _callback = callback;

        public C2paToolProcessResult Run(
            C2paToolProcessRequest request,
            CancellationToken cancellationToken = default) {
            _callback();
            return new C2paToolProcessResult(
                0,
                "{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[]}",
                string.Empty);
        }
    }
}
