using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows.Tests;

public sealed partial class OfficeProvenanceWorkflowTests {
    [Theory]
    [InlineData("null-verifier")]
    [InlineData("inconsistent-verifier")]
    [InlineData("null-detector")]
    [InlineData("inconsistent-detector")]
    public async Task ProviderContractFailuresUseTheExecutionFailureContract(string scenario) {
        using var scope = new TempScope();
        string input = scope.Write("asset.txt", "provider contract input");
        IOfficeProvenanceVerifier? verifier = scenario switch {
            "null-verifier" => new NullResultVerifier(),
            "inconsistent-verifier" => new InconsistentResultVerifier(),
            _ => null
        };
        IOfficeProvenanceSignalDetector[] detectors = scenario switch {
            "null-detector" => [new NullResultDetector()],
            "inconsistent-detector" => [new InconsistentResultDetector()],
            _ => []
        };

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(verifier, detectors)
            .RunProvenanceAsync(new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.False(result.Succeeded);
        Assert.Equal(OfficeWorkflowFailureKind.OperationFailed, result.FailureKind);
        Assert.Contains("returned", result.Summary, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class NullResultVerifier : IOfficeProvenanceVerifier {
        public string Name => "null-verifier";
        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) => null!;
    }

    private sealed class InconsistentResultVerifier : IOfficeProvenanceVerifier {
        public string Name => "expected-verifier";
        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) =>
            new(OfficeProvenanceVerificationStatus.Valid, "different-verifier", []);
    }

    private sealed class NullResultDetector : IOfficeProvenanceSignalDetector {
        public string Name => "null-detector";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;
        public OfficeProvenanceSignalResult Detect(string filePath) => null!;
    }

    private sealed class InconsistentResultDetector : IOfficeProvenanceSignalDetector {
        public string Name => "expected-detector";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;
        public OfficeProvenanceSignalResult Detect(string filePath) =>
            new("different-detector", SignalKind, OfficeProvenanceSignalStatus.NotDetected);
    }
}
