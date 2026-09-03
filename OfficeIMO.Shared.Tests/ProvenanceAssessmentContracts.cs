using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceAssessmentContracts {
    [Fact]
    public void SnapshotUsesPrivateUnixPermissionsAndRemovesItsPayload() {
#if NET8_0_OR_GREATER
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string source = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(source, "sensitive", new UTF8Encoding(false));
        string snapshotPath;
        string snapshotDirectory;
        try {
            using (OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.Capture(source, 1024)) {
                snapshotPath = snapshot.FilePath;
                snapshotDirectory = Path.GetDirectoryName(snapshotPath)!;
                Assert.Equal(UnixFileMode.UserRead, File.GetUnixFileMode(snapshotPath));
                Assert.Equal(
                    UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute,
                    File.GetUnixFileMode(snapshotDirectory));
                Assert.Equal(Path.GetFileName(source), Path.GetFileName(snapshotPath));
            }
            Assert.False(File.Exists(snapshotPath));
            Assert.False(Directory.Exists(snapshotDirectory));
        } finally {
            File.Delete(source);
    }
#endif
    }

    [Fact]
    public void SnapshotCleanupCanBeRetriedAfterAWindowsSharingViolation() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string input = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(input, "snapshot");
        OfficeProvenanceFileSnapshot? snapshot = null;
        try {
            snapshot = OfficeProvenanceFileSnapshot.Capture(input, maximumBytes: 1024);
            string snapshotPath = snapshot.FilePath;
            using (var blocker = new FileStream(
                       snapshotPath,
                       FileMode.Open,
                       FileAccess.Read,
                       FileShare.Read)) {
                IOException exception = Assert.Throws<IOException>(() => snapshot.Dispose());
                Assert.Contains(snapshotPath, exception.Message, StringComparison.Ordinal);
                Assert.True(File.Exists(snapshotPath));
            }

            snapshot.Dispose();
            Assert.False(File.Exists(snapshotPath));
            snapshot = null;
        } finally {
            snapshot?.Dispose();
            if (File.Exists(input)) File.Delete(input);
        }
    }

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
    public void AssessmentInspectFileUsesOneImmutableSnapshotForEveryProvider() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "original marker\u200B", new UTF8Encoding(false));
        var detector = new ReplacingDetector(path);
        try {
            OfficeProvenanceAssessmentReport report = OfficeProvenanceAssessment.InspectFile(
                path,
                signalDetectors: [detector]);

            Assert.Equal(OfficeTextIntegrityFindingKind.ZeroWidthSpace, Assert.Single(report.TextIntegrity!.Findings).Kind);
            Assert.Equal(OfficeProvenanceSignalStatus.Detected, Assert.Single(report.ProviderSignals).Status);
            Assert.NotNull(detector.ObservedPath);
            Assert.NotEqual(Path.GetFullPath(path), detector.ObservedPath);
            Assert.False(File.Exists(detector.ObservedPath));
            Assert.Equal("replacement", File.ReadAllText(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void AssessmentInspectFileReportsTheOriginalTextLocation() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".html");
        File.WriteAllText(path, "<!doctype html><html><body>review\u200Bthis</body></html>", new UTF8Encoding(false));
        try {
            OfficeProvenanceAssessmentReport report = OfficeProvenanceAssessment.InspectFile(path);

            OfficeTextIntegrityFinding finding = Assert.Single(report.TextIntegrity!.Findings);
            Assert.Equal(Path.GetFullPath(path), finding.Location);
            Assert.DoesNotContain("officeimo-provenance-", finding.Location, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void SnapshotCapturesRelativeExternalManifestDependenciesFromFormatReports() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "page.html");
        string sidecar = Path.Combine(directory, "manifests", "c2pa", "claim.c2pa");
        Directory.CreateDirectory(Path.GetDirectoryName(sidecar)!);
        File.WriteAllText(
            path,
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"manifests/c2pa/claim.c2pa\"></head><body>body</body></html>",
            new UTF8Encoding(false));
        File.WriteAllText(sidecar, "immutable claim", new UTF8Encoding(false));
        string? snapshotDirectory = null;
        try {
            OfficeProvenanceReport report = HtmlProvenance.InspectFile(path);
            using (OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.Capture(path, 4096)) {
                snapshotDirectory = Path.GetDirectoryName(snapshot.FilePath)!;
                snapshot.CaptureExternalManifestDependencies(path, report, 4096, 4096);

                Assert.Equal(
                    "immutable claim",
                    File.ReadAllText(Path.Combine(snapshotDirectory, "manifests", "c2pa", "claim.c2pa")));
            }
            Assert.False(Directory.Exists(snapshotDirectory));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void CoreAssessmentPreservesRelativeHtmlManifestForVerifier() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(directory, "claims"));
        string path = Path.Combine(directory, "page.html");
        File.WriteAllText(
            path,
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claims/claim.c2pa\"></head><body>body</body></html>",
            new UTF8Encoding(false));
        File.WriteAllText(Path.Combine(directory, "claims", "claim.c2pa"), "immutable claim", new UTF8Encoding(false));
        try {
            OfficeProvenanceAssessmentReport report = OfficeProvenanceAssessment.InspectFile(
                path,
                verifier: new RelativeSidecarVerifier("claims/claim.c2pa", "immutable claim"));

            OfficeProvenanceEvidence evidence = Assert.Single(report.Structural.Evidence);
            Assert.Equal(OfficeProvenanceCarrierKind.C2paExternalManifest, evidence.Carrier);
            Assert.True(evidence.IsStructurallyValid);
            Assert.Equal(OfficeProvenanceVerificationStatus.Valid, report.Verification!.Status);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void CoreHtmlInspectionResolvesRelativeBaseForManifestSnapshot() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(directory, "sub"));
        string path = Path.Combine(directory, "page.html");
        File.WriteAllText(
            path,
            "<!doctype html><html><head><base href=\"sub/\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>body</body></html>",
            new UTF8Encoding(false));
        File.WriteAllText(Path.Combine(directory, "sub", "claim.c2pa"), "based claim", new UTF8Encoding(false));
        try {
            OfficeProvenanceAssessmentReport report = OfficeProvenanceAssessment.InspectFile(
                path,
                verifier: new RelativeSidecarVerifier("sub/claim.c2pa", "based claim"));

            Assert.True(Assert.Single(report.Structural.Evidence).IsStructurallyValid);
            Assert.Equal(OfficeProvenanceVerificationStatus.Valid, report.Verification!.Status);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void CoreHtmlInspectionReportsEmbeddedManifestAssociations() {
        string manifest = Convert.ToBase64String(ProvenanceCoreContracts.CreateManifestStoreForLifecycleTests());
        byte[] html = Encoding.UTF8.GetBytes(
            $"<!doctype html><html><head><script type=\"application/c2pa\">{manifest}</script></head><body></body></html>");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(html, "page.html");

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, evidence.Carrier);
        Assert.True(evidence.IsStructurallyValid);
        Assert.True(evidence.PayloadLength > 0);
        Assert.Equal(evidence.PayloadLength, report.ExpandedInspectionBytes);
    }

    [Fact]
    public void CoreHtmlInspectionRejectsUnsafeAndCompetingManifestAssociations() {
        byte[] unsafeHtml = Encoding.UTF8.GetBytes(
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"java&#10;script:alert(1)\"></head></html>");
        byte[] competingHtml = Encoding.UTF8.GetBytes(
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"first.c2pa\"><link rel=\"c2pa-manifest\" href=\"second.c2pa\"></head></html>");

        Assert.False(Assert.Single(OfficeProvenanceInspector.Inspect(unsafeHtml, "unsafe.html").Evidence).IsStructurallyValid);
        OfficeProvenanceReport competing = OfficeProvenanceInspector.Inspect(competingHtml, "competing.html");
        Assert.Equal(2, competing.Evidence.Count);
        Assert.All(competing.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.Single(competing.Diagnostics);
    }

    [Fact]
    public void PortableSnapshotFallbackCapturesPrimaryAndRelativeDependencyFiles() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "page.html");
        File.WriteAllText(
            path,
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>portable snapshot</body></html>",
            new UTF8Encoding(false));
        File.WriteAllText(Path.Combine(directory, "claim.c2pa"), "portable claim", new UTF8Encoding(false));
        string? snapshotPath = null;
        try {
            using (OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.CapturePortable(path, 4096)) {
                snapshotPath = snapshot.FilePath;
                OfficeProvenanceReport report = OfficeProvenanceInspector.InspectFile(snapshot.FilePath);
                snapshot.CaptureExternalManifestDependencies(path, report, 4096, 4096);

                Assert.Contains("portable snapshot", File.ReadAllText(snapshot.FilePath), StringComparison.Ordinal);
                Assert.Equal("portable claim", File.ReadAllText(Path.Combine(Path.GetDirectoryName(snapshot.FilePath)!, "claim.c2pa")));
                snapshot.VerifyPrimaryFile();
                snapshot.VerifyExternalManifestDependencies();
            }
            Assert.False(File.Exists(snapshotPath));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void SnapshotDeduplicatesCaseAliasesOnACaseInsensitiveFileSystem() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "page.html");
        string sidecar = Path.Combine(directory, "Claim.c2pa");
        File.WriteAllText(path, "<!doctype html><html><body>body</body></html>", new UTF8Encoding(false));
        File.WriteAllText(sidecar, "immutable claim", new UTF8Encoding(false));
        try {
            OfficeProvenanceReport report = CreateExternalManifestReport("Claim.c2pa", "claim.c2pa");
            using OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.Capture(path, 4096);
            string snapshotDirectory = Path.GetDirectoryName(snapshot.FilePath)!;
            string differentlyCasedSnapshot = Path.Combine(
                snapshotDirectory,
                Path.GetFileName(snapshot.FilePath).ToUpperInvariant());
            if (!File.Exists(differentlyCasedSnapshot)) return;

            snapshot.CaptureExternalManifestDependencies(path, report, 4096, 4096);

            Assert.Equal(
                "immutable claim",
                File.ReadAllText(Path.Combine(snapshotDirectory, "Claim.c2pa")));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void SnapshotRejectsAUnixFifoWithoutBlockingForAWriter() {
#if NET8_0_OR_GREATER
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string fifo = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".fifo");
        try {
            Assert.Equal(0, CreateFifoUnix(fifo, 0x180));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                OfficeProvenanceFileSnapshot.Capture(fifo, maximumBytes: 1024));

            Assert.Contains("regular file", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(fifo);
        }
#endif
    }

    [Fact]
    public void AssessmentRejectsProviderMutationOfThePrimarySnapshot() {
#if NET8_0_OR_GREATER
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "original", new UTF8Encoding(false));
        try {
            Assert.Throws<InvalidDataException>(() => OfficeProvenanceAssessment.InspectFile(
                path,
                signalDetectors: [new SnapshotReplacingDetector()]));
        } finally {
            File.Delete(path);
        }
#endif
    }

    [Fact]
    public void SnapshotAppliesTheManifestLimitPerExternalDependency() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "page.html");
        File.WriteAllText(path, "<!doctype html><html><body>body</body></html>", new UTF8Encoding(false));
        File.WriteAllBytes(Path.Combine(directory, "first.c2pa"), [1, 2, 3, 4]);
        File.WriteAllBytes(Path.Combine(directory, "second.c2pa"), [5, 6, 7, 8]);
        string? snapshotDirectory = null;
        try {
            OfficeProvenanceReport report = CreateExternalManifestReport("first.c2pa", "second.c2pa");
            using (OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.Capture(path, 4096)) {
                snapshotDirectory = Path.GetDirectoryName(snapshot.FilePath)!;
                snapshot.CaptureExternalManifestDependencies(
                    path,
                    report,
                    maximumDependencyBytes: 4,
                    maximumTotalBytes: 8);

                Assert.Equal([1, 2, 3, 4], File.ReadAllBytes(Path.Combine(snapshotDirectory, "first.c2pa")));
                Assert.Equal([5, 6, 7, 8], File.ReadAllBytes(Path.Combine(snapshotDirectory, "second.c2pa")));
            }
            Assert.False(Directory.Exists(snapshotDirectory));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void SnapshotAppliesTheExpandedDataLimitAcrossExternalDependencies() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "page.html");
        File.WriteAllText(path, "<!doctype html><html><body>body</body></html>", new UTF8Encoding(false));
        File.WriteAllBytes(Path.Combine(directory, "first.c2pa"), [1, 2, 3, 4]);
        File.WriteAllBytes(Path.Combine(directory, "second.c2pa"), [5, 6, 7, 8]);
        try {
            OfficeProvenanceReport report = CreateExternalManifestReport("first.c2pa", "second.c2pa");
            using OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.Capture(path, 4096);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                snapshot.CaptureExternalManifestDependencies(
                    path,
                    report,
                    maximumDependencyBytes: 4,
                    maximumTotalBytes: 7));

            Assert.Contains("expanded-data limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void SnapshotSharesTheExpandedDataLimitWithHtmlStructuralInspection() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "page.html");
        File.WriteAllText(
            path,
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head>" +
            "<body><img src=\"data:image/png;base64,AQIDBA==\"></body></html>",
            new UTF8Encoding(false));
        File.WriteAllBytes(Path.Combine(directory, "claim.c2pa"), [1, 2, 3, 4]);
        try {
            var options = new OfficeProvenanceOptions {
                MaxAssetBytes = 4096,
                MaxManifestBytes = 4096,
                MaxExpandedContainerBytes = 7
            };
            OfficeProvenanceReport report = HtmlProvenance.InspectFile(path, options);
            Assert.Equal(4, report.ExpandedInspectionBytes);
            using OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.Capture(path, 4096);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                snapshot.CaptureExternalManifestDependencies(
                    path,
                    report,
                    maximumDependencyBytes: 4096,
                    maximumTotalBytes: options.MaxExpandedContainerBytes));

            Assert.Contains("expanded-data limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void SnapshotRejectsAnExternalManifestSymlinkThatEscapesTheSourceDirectory() {
#if NET8_0_OR_GREATER
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string outside = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".c2pa");
        string path = Path.Combine(directory, "page.html");
        string sidecar = Path.Combine(directory, "claim.c2pa");
        File.WriteAllText(path, "<!doctype html><html><body>body</body></html>", new UTF8Encoding(false));
        File.WriteAllText(outside, "outside claim", new UTF8Encoding(false));
        try {
            try {
                File.CreateSymbolicLink(sidecar, outside);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or PlatformNotSupportedException) {
                return;
            }
            OfficeProvenanceReport report = CreateExternalManifestReport("claim.c2pa");
            using OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.Capture(path, 4096);

            Assert.Throws<InvalidDataException>(() => snapshot.CaptureExternalManifestDependencies(
                path,
                report,
                maximumDependencyBytes: 4096,
                maximumTotalBytes: 4096));
        } finally {
            File.Delete(sidecar);
            File.Delete(outside);
            Directory.Delete(directory, recursive: true);
        }
#endif
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

    private static OfficeProvenanceReport CreateExternalManifestReport(params string[] references) =>
        new OfficeProvenanceReport(
            OfficeProvenanceAssetFormat.Html,
            references.Select(reference => new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paExternalManifest,
                "html:link[rel=c2pa-manifest]",
                isStructurallyValid: true,
                value: reference)).ToArray());

    private sealed class StubVerifier : IOfficeProvenanceVerifier {
        public string Name => "stub-verifier";
        internal OfficeProvenanceVerificationOptions? Options { get; private set; }
        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) =>
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
        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) =>
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

    private sealed class ReplacingDetector : IOfficeProvenanceSignalDetector {
        private readonly string _originalPath;

        internal ReplacingDetector(string originalPath) => _originalPath = originalPath;

        public string Name => "replacing";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;
        internal string? ObservedPath { get; private set; }

        public OfficeProvenanceSignalResult Detect(string filePath) {
            ObservedPath = Path.GetFullPath(filePath);
            File.WriteAllText(_originalPath, "replacement", new UTF8Encoding(false));
            bool detected = File.ReadAllText(filePath).Contains("original marker", StringComparison.Ordinal);
            return new OfficeProvenanceSignalResult(
                Name,
                SignalKind,
                detected ? OfficeProvenanceSignalStatus.Detected : OfficeProvenanceSignalStatus.NotDetected);
        }
    }

    private sealed class RelativeSidecarVerifier : IOfficeProvenanceVerifier {
        private readonly string _relativePath;
        private readonly string _expectedContent;

        internal RelativeSidecarVerifier(string relativePath, string expectedContent) {
            _relativePath = relativePath;
            _expectedContent = expectedContent;
        }

        public string Name => "relative-sidecar";

        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) {
            string sidecarPath = Path.Combine(Path.GetDirectoryName(filePath)!, _relativePath);
            bool valid = File.Exists(sidecarPath) && File.ReadAllText(sidecarPath) == _expectedContent;
            return new OfficeProvenanceVerificationResult(
                valid ? OfficeProvenanceVerificationStatus.Valid : OfficeProvenanceVerificationStatus.Invalid,
                Name,
                Array.Empty<string>());
        }
    }

#if NET8_0_OR_GREATER
    [DllImport("libc", EntryPoint = "mkfifo", SetLastError = true)]
    private static extern int CreateFifoUnix(string path, uint mode);

    private sealed class SnapshotReplacingDetector : IOfficeProvenanceSignalDetector {
        public string Name => "snapshot-replacing";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;

        public OfficeProvenanceSignalResult Detect(string filePath) {
            string replacement = filePath + ".replacement";
            File.WriteAllText(replacement, "replacement", new UTF8Encoding(false));
            File.Move(replacement, filePath, overwrite: true);
            return new OfficeProvenanceSignalResult(
                Name,
                SignalKind,
                OfficeProvenanceSignalStatus.Detected);
        }
    }
#endif

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
