using System.Threading;
using OfficeIMO.Provenance.C2pa;
using Xunit;

namespace OfficeIMO.Provenance.C2pa.Tests;

public sealed class C2paToolAvailabilityTests {
    [Theory]
    [InlineData(0, "c2patool 0.27.22", true)]
    [InlineData(0, "another-tool 1.0", false)]
    [InlineData(1, "c2patool 0.27.22", false)]
    public void AvailabilityRequiresRecognizedSuccessfulVersion(int exitCode, string output, bool expected) {
        var runner = new VersionRunner(exitCode, output);
        var result = new C2paToolProvenanceVerifier("explicit-tool", runner).CheckAvailability();
        Assert.Equal(expected, result.Available);
        Assert.Equal("explicit-tool", result.ExecutablePath);
        Assert.Equal(new[] { "--version" }, runner.Request!.Arguments);
        Assert.True(runner.Request.MaximumOutputBytes <= 4096);
    }

    [Fact]
    public void CancellationPreventsVersionProbe() {
        var runner = new VersionRunner(0, "c2patool 0.27.22");
        Assert.Throws<OperationCanceledException>(() => new C2paToolProvenanceVerifier("tool", runner)
            .CheckAvailability(cancellationToken: new CancellationToken(true)));
        Assert.Null(runner.Request);
    }

    private sealed class VersionRunner(int exitCode, string output) : IC2paToolProcessRunner {
        internal C2paToolProcessRequest? Request;
        public C2paToolProcessResult Run(C2paToolProcessRequest request, CancellationToken cancellationToken = default) {
            Request = request;
            return new C2paToolProcessResult(exitCode, output, "");
        }
    }
}
