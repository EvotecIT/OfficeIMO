using System.Text;
using System.Threading;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceAssessmentContracts {
    [Fact]
    public void AssessmentKeepsStructuralVerificationTextAndProviderEvidenceDistinct() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".py");
        File.WriteAllText(path, "# text\u200B\n", new UTF8Encoding(false));
        try {
            var verifier = new StubVerifier();
            var detected = new StubDetector(
                "media-watermark",
                OfficeProvenanceSignalKind.DurableMediaWatermark,
                OfficeProvenanceSignalStatus.Detected);
            var unavailable = new StubDetector(
                "text-watermark",
                OfficeProvenanceSignalKind.StatisticalTextWatermark,
                OfficeProvenanceSignalStatus.ProviderUnavailable);

            var options = new OfficeProvenanceAssessmentOptions();
            options.Verification.IncludeRawReport = true;
            OfficeProvenanceAssessmentReport report = OfficeProvenanceAssessment.InspectFile(
                path,
                options,
                verifier: verifier,
                signalDetectors: new IOfficeProvenanceSignalDetector[] { detected, unavailable });

            Assert.Equal(OfficeProvenanceAssetFormat.StructuredText, report.Structural.Format);
            Assert.Empty(report.Structural.Evidence);
            Assert.True(report.HasVerifiedContentCredential);
            Assert.True(report.HasDetectedProviderSignal);
            Assert.Equal(OfficeTextIntegrityFindingKind.ZeroWidthSpace, Assert.Single(report.TextIntegrity!.Findings).Kind);
            Assert.Same(options.Verification, verifier.Options);
            Assert.Collection(report.ProviderSignals,
                item => Assert.Equal(OfficeProvenanceSignalStatus.Detected, item.Status),
                item => Assert.Equal(OfficeProvenanceSignalStatus.ProviderUnavailable, item.Status));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void AssessmentRejectsDetectorIdentityDrift() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "text", new UTF8Encoding(false));
        try {
            var detector = new InconsistentDetector();

            Assert.Throws<InvalidDataException>(() => OfficeProvenanceAssessment.InspectFile(
                path,
                signalDetectors: new[] { detector }));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void AssessmentRejectsVerifierIdentityDrift() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "text", new UTF8Encoding(false));
        try {
            Assert.Throws<InvalidDataException>(() => OfficeProvenanceAssessment.InspectFile(
                path,
                verifier: new InconsistentVerifier()));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void AssessmentComposesAnExistingStructuralReportWithoutChangingItsOwnerEvidence() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "text\u200B", new UTF8Encoding(false));
        try {
            OfficeProvenanceReport structural = OfficeProvenanceInspector.InspectFile(path);

            OfficeProvenanceAssessmentReport report = OfficeProvenanceAssessment.AssessFile(path, structural);

            Assert.Same(structural, report.Structural);
            Assert.Equal(OfficeTextIntegrityFindingKind.ZeroWidthSpace, Assert.Single(report.TextIntegrity!.Findings).Kind);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void AssessmentObservesCancellationRaisedByASynchronousProvider() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "text", new UTF8Encoding(false));
        using var cancellation = new CancellationTokenSource();
        try {
            OfficeProvenanceReport structural = OfficeProvenanceInspector.InspectFile(path);

            Assert.Throws<OperationCanceledException>(() => OfficeProvenanceAssessment.AssessFile(
                path,
                structural,
                signalDetectors: [new CancellingDetector(cancellation)],
                cancellationToken: cancellation.Token));
        } finally {
            File.Delete(path);
        }
    }

    private sealed class StubVerifier : IOfficeProvenanceVerifier {
        public string Name => "stub-verifier";
        internal OfficeProvenanceVerificationOptions? Options { get; private set; }
        public OfficeProvenanceVerificationResult Verify(string filePath, OfficeProvenanceVerificationOptions? options = null) =>
            Create(options);

        private OfficeProvenanceVerificationResult Create(OfficeProvenanceVerificationOptions? options) {
            Options = options;
            return new OfficeProvenanceVerificationResult(
                OfficeProvenanceVerificationStatus.Valid,
                Name,
                Array.Empty<string>());
        }
    }

    private sealed class InconsistentVerifier : IOfficeProvenanceVerifier {
        public string Name => "expected";
        public OfficeProvenanceVerificationResult Verify(string filePath, OfficeProvenanceVerificationOptions? options = null) =>
            new OfficeProvenanceVerificationResult(
                OfficeProvenanceVerificationStatus.Valid,
                "different",
                Array.Empty<string>());
    }

    private sealed class StubDetector : IOfficeProvenanceSignalDetector {
        private readonly OfficeProvenanceSignalStatus _status;
        internal StubDetector(string name, OfficeProvenanceSignalKind kind, OfficeProvenanceSignalStatus status) {
            Name = name;
            SignalKind = kind;
            _status = status;
        }
        public string Name { get; }
        public OfficeProvenanceSignalKind SignalKind { get; }
        public OfficeProvenanceSignalResult Detect(string filePath) =>
            new OfficeProvenanceSignalResult(Name, SignalKind, _status);
    }

    private sealed class InconsistentDetector : IOfficeProvenanceSignalDetector {
        public string Name => "expected";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DurableMediaWatermark;
        public OfficeProvenanceSignalResult Detect(string filePath) =>
            new OfficeProvenanceSignalResult(
                "different",
                OfficeProvenanceSignalKind.DurableMediaWatermark,
                OfficeProvenanceSignalStatus.NotDetected);
    }

    private sealed class CancellingDetector : IOfficeProvenanceSignalDetector {
        private readonly CancellationTokenSource _cancellation;

        internal CancellingDetector(CancellationTokenSource cancellation) => _cancellation = cancellation;

        public string Name => "cancelling";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;

        public OfficeProvenanceSignalResult Detect(string filePath) {
            _cancellation.Cancel();
            return new OfficeProvenanceSignalResult(Name, SignalKind, OfficeProvenanceSignalStatus.NotDetected);
        }
    }
}
