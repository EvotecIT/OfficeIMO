using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;
using static OfficeIMO.Workflows.OfficeProvenanceWorkflowAdapter;

namespace OfficeIMO.Workflows;

/// <summary>Runs cross-format provenance workflows through the owning format packages.</summary>
public sealed partial class OfficeWorkflowRunner : IOfficeProvenanceWorkflowRunner {
    private readonly IOfficeProvenanceVerifier? _provenanceVerifier;
    private readonly IReadOnlyList<IOfficeProvenanceSignalDetector> _provenanceSignalDetectors;

    /// <summary>Creates a workflow runner without optional provenance providers.</summary>
    public OfficeWorkflowRunner() : this(null, null) { }

    /// <summary>Creates a workflow runner with optional cryptographic and provider-specific provenance services.</summary>
    public OfficeWorkflowRunner(
        IOfficeProvenanceVerifier? provenanceVerifier,
        IEnumerable<IOfficeProvenanceSignalDetector>? provenanceSignalDetectors = null) {
        _provenanceVerifier = provenanceVerifier;
        _provenanceSignalDetectors = (provenanceSignalDetectors ?? Array.Empty<IOfficeProvenanceSignalDetector>())
            .Select(detector => detector ?? throw new ArgumentException(
                "Signal detector collections cannot contain null entries.", nameof(provenanceSignalDetectors)))
            .ToArray();
    }

    /// <summary>Runs one provenance request with bounded input and atomic removal publication.</summary>
    public async Task<OfficeProvenanceWorkflowResult> RunProvenanceAsync(
        OfficeProvenanceWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        var stopwatch = Stopwatch.StartNew();
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        ValidatedProvenanceRequest? validated = null;
        string? stagingPath = null;
        OfficeProvenanceFileSnapshot? inputSnapshot = null;
        long inputBytes = 0;
        WorkflowFailureStage failureStage = WorkflowFailureStage.Validation;

        try {
            cancellationToken.ThrowIfCancellationRequested();
            validated = ValidateProvenanceRequest(request);
            string ownerPackage = GetPackage(validated.Owner);
            Report(progress, validated.Id, "validate", "Validating provenance input and limits", 0.05D);
            cancellationToken.ThrowIfCancellationRequested();
            failureStage = WorkflowFailureStage.Input;
            inputBytes = new FileInfo(validated.InputPath).Length;
            long operationInputLimit = GetOperationInputLimit(validated);
            EnforceInputLimit(validated.InputPath, inputBytes, operationInputLimit);
            failureStage = WorkflowFailureStage.Snapshot;
            inputSnapshot = OfficeProvenanceFileSnapshot.Capture(
                validated.InputPath,
                operationInputLimit,
                cancellationToken);
            inputSnapshot.SealForProviderAccess();
            string operationInputPath = inputSnapshot.FilePath;
            inputBytes = inputSnapshot.Length;
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                validated.Operation switch {
                    OfficeProvenanceWorkflowOperation.Inspect => "InspectionSnapshot",
                    OfficeProvenanceWorkflowOperation.Assess => "AssessmentSnapshot",
                    _ => "RemovalSnapshot"
                },
                validated.Operation switch {
                    OfficeProvenanceWorkflowOperation.Inspect =>
                        "Structural provenance was inspected from one bounded immutable input snapshot.",
                    OfficeProvenanceWorkflowOperation.Assess =>
                        "Structural, text-integrity, verification, and provider evidence were collected from one bounded immutable input snapshot.",
                    _ => "Removal preflight and mutation used one bounded immutable input snapshot."
                },
                stage: "validate"));

            Report(progress, validated.Id, "inspect", "Inspecting through " + ownerPackage, 0.2D);
            failureStage = WorkflowFailureStage.Operation;
            OfficeProvenanceOptions inspectionOptions = validated.Operation switch {
                OfficeProvenanceWorkflowOperation.Assess => validated.Assessment.Structural,
                OfficeProvenanceWorkflowOperation.Remove => validated.RemovalInputInspection,
                _ => validated.Inspection
            };
            OfficeProvenanceReport structural = await Task.Run(
                () => OfficeProvenanceWorkflowAdapter.Inspect(
                    validated.Owner,
                    operationInputPath,
                    inspectionOptions,
                    validated.InputPath,
                    cancellationToken),
                cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            ProvenanceOwner refinedOwner = Refine(validated.Owner, structural.Format);
            if (structural.Format == OfficeProvenanceAssetFormat.Unknown) {
                throw new NotSupportedException("The input is not a supported provenance asset.");
            }
            if (refinedOwner != validated.Owner) {
                structural = await Task.Run(
                    () => OfficeProvenanceWorkflowAdapter.Inspect(
                        refinedOwner,
                        operationInputPath,
                        inspectionOptions,
                        validated.InputPath,
                        cancellationToken),
                    cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
            }
            validated = validated with { Owner = refinedOwner };
            ownerPackage = GetPackage(refinedOwner);

            if (validated.Operation == OfficeProvenanceWorkflowOperation.Inspect) {
                inputSnapshot.VerifyPrimaryFile(cancellationToken);
                inputSnapshot.Dispose();
                inputSnapshot = null;
                Report(progress, validated.Id, "complete", "Structural provenance report is ready", 1D);
                return CreateProvenanceResult(
                    validated, OfficeWorkflowStatus.Completed, OfficeWorkflowFailureKind.None,
                    ownerPackage, null, inputBytes, 0, stopwatch.Elapsed,
                    DescribeInspection(structural), diagnostics, inspection: structural);
            }

            if (validated.Operation == OfficeProvenanceWorkflowOperation.Assess) {
                Report(progress, validated.Id, "assess", "Collecting optional verification and signal evidence", 0.55D);
                Encoding? textEncoding = validated.Assessment.InspectTextIntegrity && IsTextLike(structural.Format)
                    ? OfficeProvenanceWorkflowAdapter.ResolveTextEncoding(
                        refinedOwner,
                        structural.Format,
                        operationInputPath,
                        validated.Assessment.TextIntegrity.MaxEncodedBytes,
                        cancellationToken)
                    : null;
                bool hasExternalProviders = _provenanceVerifier != null || _provenanceSignalDetectors.Count != 0;
                if (hasExternalProviders) {
                    inputSnapshot!.CaptureExternalManifestDependencies(
                        validated.InputPath,
                        structural,
                        validated.Assessment.Structural.MaxManifestBytes,
                        validated.Assessment.Structural.MaxExpandedContainerBytes,
                        cancellationToken);
                }
                OfficeProvenanceAssessmentReport assessment = await Task.Run(
                    () => OfficeProvenanceAssessment.AssessSnapshotFile(
                        operationInputPath,
                        validated.InputPath,
                        structural,
                        validated.Assessment,
                        _provenanceVerifier,
                        _provenanceSignalDetectors,
                        cancellationToken,
                        textEncoding),
                    cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                inputSnapshot!.VerifyPrimaryFile(cancellationToken);
                if (hasExternalProviders) inputSnapshot!.VerifyExternalManifestDependencies(cancellationToken);
                inputSnapshot!.Dispose();
                inputSnapshot = null;
                Report(progress, validated.Id, "complete", "Provenance assessment is ready", 1D);
                return CreateProvenanceResult(
                    validated, OfficeWorkflowStatus.Completed, OfficeWorkflowFailureKind.None,
                    ownerPackage, null, inputBytes, 0, stopwatch.Elapsed,
                    DescribeAssessment(assessment), diagnostics, assessment: assessment);
            }

            if (refinedOwner == ProvenanceOwner.Core && !SupportsCoreRemoval(structural.Format)) {
                throw new NotSupportedException(
                    "Inspection is supported, but no safe mutation owner is registered for this asset format.");
            }

            long remainingExpandedBytes = validated.RemovalInputInspection.MaxExpandedContainerBytes;
            ConsumeExpandedProcessingBytes(ref remainingExpandedBytes, structural.ExpandedInspectionBytes);

            failureStage = WorkflowFailureStage.Output;
            string outputDirectory = Path.GetDirectoryName(validated.OutputPath!)!;
            Directory.CreateDirectory(outputDirectory);
            stagingPath = Path.Combine(
                outputDirectory,
                "." + Path.GetFileNameWithoutExtension(validated.OutputPath) + "." +
                Guid.NewGuid().ToString("N") + Path.GetExtension(validated.OutputPath));
            Report(progress, validated.Id, "remove", "Removing selected carriers through " + ownerPackage, 0.48D);
            failureStage = WorkflowFailureStage.Operation;
            OfficeProvenanceRemovalResult removal;
            try {
                OfficeProvenanceRemovalOptions removalOptions = CloneRemovalOptions(
                    validated.Removal,
                    validated.Removal.Limits.MaxAssetBytes,
                    validated.Removal.EffectiveMaxOutputBytes);
                removalOptions.Limits.MaxExpandedContainerBytes = Math.Max(1L, remainingExpandedBytes);
                removal = await Task.Run(
                    () => OfficeProvenanceWorkflowAdapter.Remove(
                        refinedOwner, operationInputPath, stagingPath, removalOptions, cancellationToken),
                    cancellationToken).ConfigureAwait(false);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                failureStage = WorkflowFailureStage.Output;
                throw;
            }
            cancellationToken.ThrowIfCancellationRequested();
            ConsumeExpandedProcessingBytes(ref remainingExpandedBytes, removal.Before.ExpandedInspectionBytes);
            ConsumeExpandedProcessingBytes(ref remainingExpandedBytes, removal.After.ExpandedInspectionBytes);
            inputSnapshot!.VerifyPrimaryFile(cancellationToken);

            failureStage = WorkflowFailureStage.Output;
            long stagedBytes = new FileInfo(stagingPath).Length;
            if (stagedBytes > validated.Limits.MaximumOutputBytes) {
                throw new InvalidOperationException(
                    $"Generated artifact is {stagedBytes:N0} bytes, above the configured {validated.Limits.MaximumOutputBytes:N0}-byte limit.");
            }

            Report(progress, validated.Id, "validate-output", "Reopening the staged artifact through " + ownerPackage, 0.72D);
            using StagedArtifactFingerprint stagedFingerprint = StagedArtifactFingerprint.CaptureExpected(
                stagingPath,
                validated.Limits.MaximumOutputBytes,
                removal.DataLength,
                removal.ComputeDataSha256(cancellationToken),
                cancellationToken);
            OfficeProvenanceOptions outputInspectionOptions = CloneInspectionOptions(
                validated.RemovalOutputInspection,
                validated.RemovalOutputInspection.MaxAssetBytes);
            outputInspectionOptions.MaxExpandedContainerBytes = Math.Max(1L, remainingExpandedBytes);
            OfficeProvenanceReport reopened = await Task.Run(
                () => OfficeProvenanceWorkflowAdapter.Inspect(
                    refinedOwner,
                    stagingPath,
                    outputInspectionOptions,
                    validated.OutputPath,
                    cancellationToken),
                cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            ConsumeExpandedProcessingBytes(ref remainingExpandedBytes, reopened.ExpandedInspectionBytes);
            EnsureEquivalent(removal.After, reopened);
            stagedFingerprint.VerifyStagingPath(
                stagingPath,
                validated.Limits.MaximumOutputBytes,
                cancellationToken);
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "ProvenanceOutputReopened",
                "The staged artifact was reopened through its owning format API and matched the in-memory removal report.",
                stage: "validate-output",
                details: new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["ownerPackage"] = ownerPackage,
                    ["remainingCarriers"] = reopened.Evidence.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["stagedBytes"] = stagedBytes.ToString(System.Globalization.CultureInfo.InvariantCulture)
            }));
            cancellationToken.ThrowIfCancellationRequested();

            failureStage = WorkflowFailureStage.Output;
            string ownedStagingPath = stagingPath;
            stagingPath = null;
            string publishedPath = PublishVerified(
                ownedStagingPath,
                validated.OutputPath!,
                validated.ConflictPolicy,
                validated.BatchBlockedOutputIdentities,
                validated.BatchOwnReservedOutputIdentity,
                stagedFingerprint,
                validated.ConflictPolicy == OfficeWorkflowConflictPolicy.Replace &&
                OfficeWorkflowPathIdentity.AreEquivalent(validated.InputPath, validated.OutputPath!)
                    ? inputSnapshot
                    : null,
                validated.Limits.MaximumOutputBytes,
                cancellationToken,
                beforePublish: () => Report(
                    progress,
                    validated.Id,
                    "publish",
                    "Publishing the verified provenance artifact",
                    0.9D),
                beforeCommitFinalized: path => {
                    Report(progress, validated.Id, "finalize", "Finalizing the verified provenance artifact", 0.98D);
                    stagedFingerprint.VerifyPublishedPath(path, validated.Limits.MaximumOutputBytes, cancellationToken);
                });
            inputSnapshot.Dispose();
            inputSnapshot = null;
            long outputBytes = stagedFingerprint.Length;
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "AtomicPublication",
                "The verified artifact was staged in the destination directory, identity-pinned, and atomically published.",
                stage: "publish"));
            Report(progress, validated.Id, "complete", "Provenance removal completed", 1D);
            return CreateProvenanceResult(
                validated, OfficeWorkflowStatus.Completed, OfficeWorkflowFailureKind.None,
                ownerPackage, publishedPath, inputBytes, outputBytes, stopwatch.Elapsed,
                DescribeRemoval(removal, reopened), diagnostics,
                before: removal.Before,
                after: reopened,
                changes: removal.Changes,
                wasReserialized: removal.WasReserialized,
                wereInvalidatedSignaturesRemoved: removal.WereInvalidatedSignaturesRemoved);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            TryDisposeSnapshot(ref inputSnapshot, diagnostics);
            bool stagingCleaned = TryCleanupStaging(ref stagingPath, diagnostics);
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "Cancelled",
                stagingCleaned
                    ? "The provenance workflow was cancelled before publication; no staged artifact was retained."
                    : "The provenance workflow was cancelled before publication, but staging cleanup failed; the retained path is reported in diagnostics.",
                OfficeWorkflowDiagnosticSeverity.Information,
                "cancel"));
            return new OfficeProvenanceWorkflowResult(
                validated?.Id ?? request.Id,
                validated?.Operation ?? request.Operation,
                OfficeWorkflowStatus.Cancelled,
                OfficeWorkflowFailureKind.None,
                GetResultPackage(validated, request),
                null,
                inputBytes,
                0,
                stopwatch.Elapsed,
                "Cancelled",
                diagnostics);
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            TryDisposeSnapshot(ref inputSnapshot, diagnostics);
            TryCleanupStaging(ref stagingPath, diagnostics);
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "ProvenanceWorkflowFailed",
                exception.Message,
                OfficeWorkflowDiagnosticSeverity.Error,
                GetDiagnosticStage(failureStage),
                new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["exceptionType"] = exception.GetType().Name
                }));
            return new OfficeProvenanceWorkflowResult(
                validated?.Id ?? request.Id,
                validated?.Operation ?? request.Operation,
                OfficeWorkflowStatus.Failed,
                ClassifyFailure(exception, failureStage),
                GetResultPackage(validated, request),
                null,
                inputBytes,
                0,
                stopwatch.Elapsed,
                "Provenance workflow failed: " + exception.Message,
                diagnostics);
        } finally {
            try {
                inputSnapshot?.Dispose();
            } catch (Exception) when (inputSnapshot is not null) {
                // Catch paths already record cleanup failures. Never let one skip staged-output cleanup.
            } finally {
                if (stagingPath is not null) TryDelete(stagingPath);
            }
        }
    }

    private static string GetResultPackage(
        ValidatedProvenanceRequest? validated,
        OfficeProvenanceWorkflowRequest request) {
        if (validated is not null) return GetPackage(validated.Owner);
        try {
            return GetPackage(ResolveByPath(request.InputPath));
        } catch (Exception exception) when (exception is ArgumentException or NotSupportedException) {
            return "OfficeIMO.Core";
        }
    }

    private static void TryDisposeSnapshot(
        ref OfficeProvenanceFileSnapshot? snapshot,
        ICollection<OfficeWorkflowDiagnostic> diagnostics) {
        if (snapshot is null) return;
        try {
            snapshot.Dispose();
            snapshot = null;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "ProvenanceSnapshotCleanupFailed",
                exception.Message,
                OfficeWorkflowDiagnosticSeverity.Error,
                "cleanup",
                new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["exceptionType"] = exception.GetType().Name
                }));
        }
    }

    private static bool TryCleanupStaging(
        ref string? stagingPath,
        ICollection<OfficeWorkflowDiagnostic> diagnostics) {
        if (stagingPath is null) return true;
        string retainedPath = stagingPath;
        try {
            File.Delete(retainedPath);
            stagingPath = null;
            return true;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            stagingPath = null;
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "ProvenanceStagingCleanupFailed",
                $"The staged provenance artifact could not be removed; '{retainedPath}' is retained for operator cleanup.",
                OfficeWorkflowDiagnosticSeverity.Error,
                "cleanup",
                new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["retainedPath"] = retainedPath,
                    ["exceptionType"] = exception.GetType().Name
                }));
            return false;
        }
    }

    internal sealed class StagedArtifactFingerprint : IDisposable {
        private readonly string? _physicalIdentity;
        private readonly bool _usesPhysicalIdentity;
        private FileStream? _lease;
        private FileStream? _publishedLease;

        private StagedArtifactFingerprint(
            long length,
            byte[] sha256,
            string? physicalIdentity,
            bool usesPhysicalIdentity,
            FileStream lease) {
            Length = length;
            Sha256 = sha256;
            _physicalIdentity = physicalIdentity;
            _usesPhysicalIdentity = usesPhysicalIdentity;
            _lease = lease;
        }

        internal long Length { get; }
        private byte[] Sha256 { get; }

        internal static StagedArtifactFingerprint Capture(
            string path,
            long maximumBytes,
            CancellationToken cancellationToken,
            string artifactDescription = "staged provenance artifact") => CaptureCore(
                path,
                maximumBytes,
                expectedLength: null,
                expectedSha256: null,
                cancellationToken,
                artifactDescription,
                OfficeWorkflowPathIdentity.SupportsPhysicalIdentity);

        /// <summary>Exercises the portable length-and-hash fingerprint path when filesystem identity is unavailable.</summary>
        internal static StagedArtifactFingerprint CapturePortable(
            string path,
            long maximumBytes,
            CancellationToken cancellationToken = default) => CaptureCore(
                path,
                maximumBytes,
                expectedLength: null,
                expectedSha256: null,
                cancellationToken,
                "staged provenance artifact",
                usesPhysicalIdentity: false);

        internal static StagedArtifactFingerprint CaptureExpected(
            string path,
            long maximumBytes,
            long expectedLength,
            byte[] expectedSha256,
            CancellationToken cancellationToken) => CaptureCore(
                path,
                maximumBytes,
                expectedLength,
                expectedSha256 ?? throw new ArgumentNullException(nameof(expectedSha256)),
                cancellationToken,
                "staged provenance artifact",
                OfficeWorkflowPathIdentity.SupportsPhysicalIdentity);

        private static StagedArtifactFingerprint CaptureCore(
            string path,
            long maximumBytes,
            long? expectedLength,
            byte[]? expectedSha256,
            CancellationToken cancellationToken,
            string artifactDescription,
            bool usesPhysicalIdentity) {
            var stream = OpenForIdentity(path);
            try {
                if (stream.Length > maximumBytes) {
                    throw OfficeProvenanceLimitException.CreateOutput(
                        $"The {artifactDescription} exceeds the configured output limit of {maximumBytes} bytes.");
                }
                byte[] sha256 = ComputeHash(stream, cancellationToken);
                if (expectedLength.HasValue &&
                    (stream.Length != expectedLength.Value ||
                     !CryptographicOperations.FixedTimeEquals(sha256, expectedSha256!))) {
                    throw new InvalidDataException(
                        "The staged provenance artifact did not match the bytes returned by its format owner.");
                }
                stream.Position = 0;
                string? physicalIdentity = usesPhysicalIdentity
                    ? OfficeWorkflowPathIdentity.GetPhysicalIdentityKey(path, stream)
                    : null;
                return new StagedArtifactFingerprint(stream.Length, sha256, physicalIdentity, usesPhysicalIdentity, stream);
            } catch {
                stream.Dispose();
                throw;
            }
        }

        internal void VerifyStagingPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            if (!MatchesPath(path, maximumBytes, cancellationToken)) {
                throw new InvalidDataException(
                    "The staged provenance artifact changed after output validation; publication was blocked.");
            }
        }

        internal bool TryPinPublishedPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            FileStream stream = OpenForIdentity(path);
            try {
                if (!MatchesStream(path, stream, maximumBytes, cancellationToken)) return false;
                _publishedLease?.Dispose();
                _publishedLease = stream;
                return true;
            } catch {
                stream.Dispose();
                throw;
            } finally {
                if (!ReferenceEquals(_publishedLease, stream)) stream.Dispose();
            }
        }

        internal void VerifyPublishedPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            if (_publishedLease is null || !MatchesPath(path, maximumBytes, cancellationToken)) {
                throw new InvalidDataException(
                    "The published provenance artifact changed before publication was finalized.");
            }
        }

        internal void ReleasePublishedLease() {
            _publishedLease?.Dispose();
            _publishedLease = null;
        }

        internal void ReleaseStagingLease() {
            _lease?.Dispose();
            _lease = null;
        }

        internal void TryDeleteMatchingPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            string quarantinePath = Path.Combine(
                Path.GetDirectoryName(Path.GetFullPath(path))!,
                ".officeimo-provenance-rollback-" + Guid.NewGuid().ToString("N") + ".tmp");
            bool moved = false;
            try {
                File.Move(path, quarantinePath, overwrite: false);
                moved = true;
                bool matches;
                using (FileStream stream = OpenForIdentity(quarantinePath)) {
                    matches = MatchesStream(quarantinePath, stream, maximumBytes, cancellationToken);
                }
                if (matches) {
                    File.Delete(quarantinePath);
                    moved = false;
                    return;
                }

                File.Move(quarantinePath, path, overwrite: false);
                moved = false;
            } catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException or IOException or UnauthorizedAccessException) {
                // Never delete a known destination pathname after a failed identity check. If a
                // different writer claimed it, retain the random quarantine rather than losing data.
            } finally {
                if (moved && File.Exists(quarantinePath)) {
                    try {
                        if (!File.Exists(path)) {
                            File.Move(quarantinePath, path, overwrite: false);
                            moved = false;
                        }
                    } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
                }
            }
        }

        public void Dispose() {
            ReleasePublishedLease();
            ReleaseStagingLease();
        }

        internal bool MatchesPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            try {
                using FileStream stream = OpenForIdentity(path);
                return MatchesStream(path, stream, maximumBytes, cancellationToken);
            } catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException) {
                return false;
            }
        }

        private bool MatchesStream(
            string path,
            FileStream stream,
            long maximumBytes,
            CancellationToken cancellationToken) {
            if (stream.Length > maximumBytes || stream.Length != Length) return false;
            if (_usesPhysicalIdentity) {
                string physicalIdentity = OfficeWorkflowPathIdentity.GetPhysicalIdentityKey(path, stream);
                if (!string.Equals(physicalIdentity, _physicalIdentity, StringComparison.Ordinal)) return false;
            }
            byte[] currentHash = ComputeHash(stream, cancellationToken);
            return CryptographicOperations.FixedTimeEquals(currentHash, Sha256);
        }

        private static FileStream OpenForIdentity(string path) => new(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            81920,
            FileOptions.SequentialScan);

        private static byte[] ComputeHash(Stream stream, CancellationToken cancellationToken) {
            stream.Position = 0;
            using IncrementalHash algorithm = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var buffer = new byte[81920];
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                algorithm.AppendData(buffer, 0, read);
            }
            return algorithm.GetHashAndReset();
        }
    }

    private static string PublishVerified(
        string stagingPath,
        string requestedPath,
        OfficeWorkflowConflictPolicy policy,
        SortedSet<string>? blockedOutputIdentities,
        string? ownReservedOutputIdentity,
        StagedArtifactFingerprint staged,
        OfficeProvenanceFileSnapshot? expectedDisplacedInput,
        long maximumBytes,
        CancellationToken cancellationToken,
        Action beforePublish,
        Action<string> beforeCommitFinalized) {
        bool published = false;
        try {
            beforePublish();
            cancellationToken.ThrowIfCancellationRequested();
            staged.VerifyStagingPath(stagingPath, maximumBytes, cancellationToken);
            staged.ReleaseStagingLease();

            string publishedPath;
            switch (policy) {
                case OfficeWorkflowConflictPolicy.Fail:
                    EnsureBatchCandidateDoesNotOverlapAnotherRequest(requestedPath);
                    File.Move(stagingPath, requestedPath, overwrite: false);
                    publishedPath = requestedPath;
                    PinAndFinalize(publishedPath);
                    break;
                case OfficeWorkflowConflictPolicy.Rename:
                    publishedPath = PublishRenamed();
                    break;
                case OfficeWorkflowConflictPolicy.Replace:
                    publishedPath = PublishReplacement();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(policy), policy, "Unsupported conflict policy.");
            }

            published = true;
            return publishedPath;
        } finally {
            if (!published) {
                staged.ReleasePublishedLease();
                staged.TryDeleteMatchingPath(stagingPath, maximumBytes, CancellationToken.None);
            }
        }

        void PinAndFinalize(string path) {
            try {
                if (!staged.TryPinPublishedPath(path, maximumBytes, cancellationToken)) {
                    throw new InvalidDataException(
                        "The staged provenance artifact changed while it was being published.");
                }
                beforeCommitFinalized(path);
            } catch {
                staged.ReleasePublishedLease();
                staged.TryDeleteMatchingPath(path, maximumBytes, CancellationToken.None);
                throw;
            }
        }

        void EnsureBatchCandidateDoesNotOverlapAnotherRequest(string path) {
            if (blockedOutputIdentities is null) return;
            string identity = OfficeWorkflowPathIdentity.Normalize(path);
            bool isOwnReservation = string.Equals(identity, ownReservedOutputIdentity, StringComparison.Ordinal);
            bool hasHierarchyCollision = TryFindAncestorOrDescendant(
                identity,
                blockedOutputIdentities,
                out string? collisionIdentity) &&
                !string.Equals(collisionIdentity, ownReservedOutputIdentity, StringComparison.Ordinal);
            if ((!isOwnReservation && blockedOutputIdentities.Contains(identity)) || hasHierarchyCollision) {
                throw new IOException(
                    "The provenance output now overlaps another batch request path and cannot be published safely.");
            }
        }

        string PublishRenamed() {
            for (int suffix = 0; suffix < 10_000; suffix++) {
                cancellationToken.ThrowIfCancellationRequested();
                string candidate = suffix == 0 ? requestedPath : AddSuffix(requestedPath, suffix);
                if (blockedOutputIdentities is not null) {
                    string identity = OfficeWorkflowPathIdentity.Normalize(candidate);
                    bool isOwnReservation = string.Equals(
                        identity,
                        ownReservedOutputIdentity,
                        StringComparison.Ordinal);
                    bool hasHierarchyCollision = TryFindAncestorOrDescendant(
                        identity,
                        blockedOutputIdentities,
                        out string? collisionIdentity) &&
                        !string.Equals(collisionIdentity, ownReservedOutputIdentity, StringComparison.Ordinal);
                    if ((!isOwnReservation && blockedOutputIdentities.Contains(identity)) ||
                        hasHierarchyCollision) continue;
                }
                try {
                    File.Move(stagingPath, candidate, overwrite: false);
                } catch (IOException) when (File.Exists(candidate) || Directory.Exists(candidate)) {
                    // Another request owns this candidate. Try the next deterministic suffix.
                    continue;
                }
                PinAndFinalize(candidate);
                return candidate;
            }
            throw new IOException("No available numbered output path could be reserved.");
        }

        string PublishReplacement() {
            EnsureBatchCandidateDoesNotOverlapAnotherRequest(requestedPath);
            if (!File.Exists(requestedPath)) {
                if (expectedDisplacedInput != null) {
                    throw new IOException(
                        "The provenance input changed while its verified replacement was being published.");
                }
                bool destinationAppeared = false;
                try {
                    File.Move(stagingPath, requestedPath, overwrite: false);
                } catch (IOException) when (File.Exists(requestedPath)) {
                    // The destination appeared during the claim. Validate and replace it below.
                    destinationAppeared = true;
                }
                if (!destinationAppeared) {
                    PinAndFinalize(requestedPath);
                    return requestedPath;
                }
            }

            EnsureBatchCandidateDoesNotOverlapAnotherRequest(requestedPath);
            if (expectedDisplacedInput != null) {
                bool inputCommitted = OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchangedAndFinalize(
                    stagingPath,
                    requestedPath,
                    backupPath => expectedDisplacedInput.MatchesCapturedSource(backupPath, cancellationToken),
                    installedPath => staged.TryPinPublishedPath(installedPath, maximumBytes, cancellationToken),
                    beforeCommitFinalized);
                if (!inputCommitted) {
                    staged.ReleasePublishedLease();
                    throw new IOException(
                        "The provenance input changed while its verified replacement was being published.");
                }
                return requestedPath;
            }

            using StagedArtifactFingerprint destination = StagedArtifactFingerprint.Capture(
                requestedPath,
                maximumBytes,
                cancellationToken,
                "existing provenance destination");
            destination.ReleaseStagingLease();
            bool committed = OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchangedAndFinalize(
                stagingPath,
                requestedPath,
                backupPath => destination.MatchesPath(backupPath, maximumBytes, cancellationToken),
                installedPath => staged.TryPinPublishedPath(installedPath, maximumBytes, cancellationToken),
                beforeCommitFinalized);
            if (!committed) {
                staged.ReleasePublishedLease();
                throw new IOException("The provenance destination changed while the verified artifact was being published.");
            }
            return requestedPath;
        }
    }

    private static ValidatedProvenanceRequest ValidateProvenanceRequest(OfficeProvenanceWorkflowRequest request) {
        if (string.IsNullOrWhiteSpace(request.Id)) throw new ArgumentException("Request id cannot be empty.", nameof(request));
        if (!Enum.IsDefined(typeof(OfficeProvenanceWorkflowOperation), request.Operation)) {
            throw new ArgumentOutOfRangeException(nameof(request), request.Operation, "Choose a supported provenance operation.");
        }
        if (!Enum.IsDefined(typeof(OfficeWorkflowConflictPolicy), request.ConflictPolicy)) {
            throw new ArgumentOutOfRangeException(nameof(request), request.ConflictPolicy, "Choose a supported output conflict policy.");
        }
        if (string.IsNullOrWhiteSpace(request.InputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(request));
        string inputPath = Path.GetFullPath(request.InputPath);
        if (!File.Exists(inputPath)) throw new FileNotFoundException("The provenance input file does not exist.", inputPath);
        OfficeWorkflowLimits limits = (request.Limits ?? throw new ArgumentException("Workflow limits cannot be null.", nameof(request))).CloneAndValidate();
        OfficeProvenanceOptions inspection = CloneInspectionOptions(
            request.Inspection ?? throw new ArgumentException("Inspection options cannot be null.", nameof(request)),
            limits);
        OfficeProvenanceAssessmentOptions assessment = CloneAssessmentOptions(
            request.Assessment ?? throw new ArgumentException("Assessment options cannot be null.", nameof(request)),
            limits);
        OfficeProvenanceRemovalOptions removalSource = request.Removal ??
            throw new ArgumentException("Removal options cannot be null.", nameof(request));
        OfficeProvenanceRemovalOptions removal = CloneRemovalOptions(
            removalSource,
            limits.MaximumInputBytes,
            limits.MaximumOutputBytes);
        OfficeProvenanceOptions removalInputInspection = CreateInspectionOptions(
            removalSource,
            limits.MaximumInputBytes);
        OfficeProvenanceOptions removalOutputInspection = CreateOutputInspectionOptions(removal);
        string? outputPath = string.IsNullOrWhiteSpace(request.OutputPath) ? null : Path.GetFullPath(request.OutputPath);

        if (request.Operation == OfficeProvenanceWorkflowOperation.Remove) {
            outputPath ??= Path.Combine(
                Path.GetDirectoryName(inputPath)!,
                Path.GetFileNameWithoutExtension(inputPath) + ".provenance-cleaned" + Path.GetExtension(inputPath));
            if (!string.Equals(Path.GetExtension(inputPath), Path.GetExtension(outputPath), StringComparison.OrdinalIgnoreCase)) {
                throw new ArgumentException("Provenance removal preserves the input format and requires the same output extension.", nameof(request));
            }
        } else if (outputPath is not null) {
            throw new ArgumentException("Inspect and assess are report-only operations and do not publish an artifact.", nameof(request));
        }

        return new ValidatedProvenanceRequest(
            request.Id,
            request.Operation,
            inputPath,
            outputPath,
            ResolveByPath(inputPath),
            request.ConflictPolicy,
            request.BatchBlockedOutputIdentities,
            request.BatchOwnReservedOutputIdentity,
            limits,
            inspection,
            assessment,
            removal,
            removalInputInspection,
            removalOutputInspection);
    }

    private static OfficeProvenanceOptions CloneInspectionOptions(OfficeProvenanceOptions source, OfficeWorkflowLimits limits) => new() {
        MaxAssetBytes = Math.Min(source.MaxAssetBytes, limits.MaximumInputBytes),
        MaxManifestBytes = Math.Min(source.MaxManifestBytes, Math.Min(source.MaxAssetBytes, limits.MaximumInputBytes)),
        MaxCarriers = source.MaxCarriers,
        MaxContainerEntries = source.MaxContainerEntries,
        MaxExpandedContainerBytes = source.MaxExpandedContainerBytes,
        ProcessEmbeddedAssets = source.ProcessEmbeddedAssets,
        MaxEmbeddedAssets = source.MaxEmbeddedAssets
    };

    private static OfficeProvenanceAssessmentOptions CloneAssessmentOptions(
        OfficeProvenanceAssessmentOptions source,
        OfficeWorkflowLimits limits) {
        var clone = new OfficeProvenanceAssessmentOptions {
            InspectTextIntegrity = source.InspectTextIntegrity
        };
        CopyInspectionOptions(source.Structural, clone.Structural, limits.MaximumInputBytes);
        clone.TextIntegrity.MaxEncodedBytes = Math.Min(source.TextIntegrity.MaxEncodedBytes, limits.MaximumInputBytes);
        clone.TextIntegrity.MaxCharacters = source.TextIntegrity.MaxCharacters;
        clone.TextIntegrity.MaxFindings = source.TextIntegrity.MaxFindings;
        clone.TextIntegrity.IgnoreLeadingByteOrderMark = source.TextIntegrity.IgnoreLeadingByteOrderMark;
        clone.TextIntegrity.IncludeTypographicSpaces = source.TextIntegrity.IncludeTypographicSpaces;
        clone.TextIntegrity.IncludeVariationSelectors = source.TextIntegrity.IncludeVariationSelectors;
        clone.Verification.Timeout = source.Verification.Timeout;
        clone.Verification.MaxReportBytes = source.Verification.MaxReportBytes;
        clone.Verification.AllowNetworkAccess = source.Verification.AllowNetworkAccess;
        clone.Verification.IncludeRawReport = source.Verification.IncludeRawReport;
        clone.Verification.TrustAnchorsPath = source.Verification.TrustAnchorsPath;
        clone.Verification.AllowedListPath = source.Verification.AllowedListPath;
        clone.Verification.TrustConfigurationPath = source.Verification.TrustConfigurationPath;
        return clone;
    }

    private static OfficeProvenanceRemovalOptions CloneRemovalOptions(
        OfficeProvenanceRemovalOptions source,
        long maximumInputBytes,
        long maximumOutputBytes) {
        var clone = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            SignatureMutationPolicy = source.SignatureMutationPolicy,
            ProcessEmbeddedAssets = source.ProcessEmbeddedAssets,
            MaxEmbeddedAssets = source.MaxEmbeddedAssets,
            MaxOutputBytes = Math.Min(source.EffectiveMaxOutputBytes, maximumOutputBytes)
        };
        CopyInspectionOptions(source.Limits, clone.Limits, maximumInputBytes);
        return clone;
    }

    private static OfficeProvenanceOptions CreateInspectionOptions(
        OfficeProvenanceRemovalOptions removal,
        long maximumAssetBytes) {
        var options = new OfficeProvenanceOptions();
        CopyInspectionOptions(removal.Limits, options, maximumAssetBytes);
        options.ProcessEmbeddedAssets = removal.ProcessEmbeddedAssets && removal.Limits.ProcessEmbeddedAssets;
        options.MaxEmbeddedAssets = Math.Min(removal.MaxEmbeddedAssets, removal.Limits.MaxEmbeddedAssets);
        return options;
    }

    private static long GetOperationInputLimit(ValidatedProvenanceRequest request) => request.Operation switch {
        OfficeProvenanceWorkflowOperation.Inspect => request.Inspection.MaxAssetBytes,
        OfficeProvenanceWorkflowOperation.Assess => request.Assessment.Structural.MaxAssetBytes,
        OfficeProvenanceWorkflowOperation.Remove => request.RemovalInputInspection.MaxAssetBytes,
        _ => request.Limits.MaximumInputBytes
    };

    private static OfficeProvenanceOptions CloneInspectionOptions(
        OfficeProvenanceOptions source,
        long maximumAssetBytes) {
        var clone = new OfficeProvenanceOptions();
        CopyInspectionOptions(source, clone, maximumAssetBytes);
        return clone;
    }

    private static void ConsumeExpandedProcessingBytes(ref long remainingBytes, long consumedBytes) {
        if (consumedBytes < 0 || consumedBytes > remainingBytes) {
            throw OfficeProvenanceLimitException.Create(
                "The provenance workflow exceeds the configured cumulative expanded-data limit.");
        }
        remainingBytes -= consumedBytes;
    }

    private static OfficeProvenanceOptions CreateOutputInspectionOptions(
        OfficeProvenanceRemovalOptions removal) {
        OfficeProvenanceOptions options = CreateInspectionOptions(
            removal,
            removal.EffectiveMaxOutputBytes);
        options.MaxAssetBytes = removal.EffectiveMaxOutputBytes;
        options.MaxManifestBytes = Math.Min(removal.Limits.MaxManifestBytes, options.MaxAssetBytes);
        return options;
    }

    private static void CopyInspectionOptions(
        OfficeProvenanceOptions source,
        OfficeProvenanceOptions destination,
        long maximumAssetBytes) {
        destination.MaxAssetBytes = Math.Min(source.MaxAssetBytes, maximumAssetBytes);
        destination.MaxManifestBytes = Math.Min(source.MaxManifestBytes, destination.MaxAssetBytes);
        destination.MaxCarriers = source.MaxCarriers;
        destination.MaxContainerEntries = source.MaxContainerEntries;
        destination.MaxExpandedContainerBytes = source.MaxExpandedContainerBytes;
        destination.ProcessEmbeddedAssets = source.ProcessEmbeddedAssets;
        destination.MaxEmbeddedAssets = source.MaxEmbeddedAssets;
    }

    private static void EnsureEquivalent(OfficeProvenanceReport expected, OfficeProvenanceReport actual) {
        bool evidenceMatches = expected.Evidence.Count == actual.Evidence.Count &&
            expected.Evidence.Zip(actual.Evidence).All(pair =>
                pair.First.Carrier == pair.Second.Carrier &&
                string.Equals(pair.First.Location, pair.Second.Location, StringComparison.Ordinal) &&
                pair.First.IsStructurallyValid == pair.Second.IsStructurallyValid &&
                pair.First.PayloadLength == pair.Second.PayloadLength &&
                string.Equals(pair.First.Value, pair.Second.Value, StringComparison.Ordinal) &&
                pair.First.DigitalSourceKind == pair.Second.DigitalSourceKind);
        if (expected.Format != actual.Format || !evidenceMatches) {
            throw new InvalidDataException("The reopened artifact did not match the provenance removal report.");
        }
    }

    private static string DescribeInspection(OfficeProvenanceReport report) =>
        $"Inspected {report.Format}; found {report.Evidence.Count:N0} structural provenance carrier(s).";

    private static string DescribeAssessment(OfficeProvenanceAssessmentReport report) =>
        $"Assessed {report.Structural.Format}; found {report.Structural.Evidence.Count:N0} structural carrier(s), " +
        $"{report.TextIntegrity?.Findings.Count ?? 0:N0} text-integrity finding(s), and " +
        $"{report.ProviderSignals.Count:N0} provider signal result(s).";

    private static string DescribeRemoval(OfficeProvenanceRemovalResult result, OfficeProvenanceReport reopened) =>
        $"Applied {result.Changes.Count:N0} provenance change(s); {reopened.Evidence.Count:N0} carrier(s) remain after verified publication.";

    private static OfficeProvenanceWorkflowResult CreateProvenanceResult(
        ValidatedProvenanceRequest request,
        OfficeWorkflowStatus status,
        OfficeWorkflowFailureKind failureKind,
        string ownerPackage,
        string? outputPath,
        long inputBytes,
        long outputBytes,
        TimeSpan duration,
        string summary,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics,
        OfficeProvenanceReport? inspection = null,
        OfficeProvenanceAssessmentReport? assessment = null,
        OfficeProvenanceReport? before = null,
        OfficeProvenanceReport? after = null,
        IReadOnlyList<OfficeProvenanceChange>? changes = null,
        bool wasReserialized = false,
        bool wereInvalidatedSignaturesRemoved = false) => new(
            request.Id,
            request.Operation,
            status,
            failureKind,
            ownerPackage,
            outputPath,
            inputBytes,
            outputBytes,
            duration,
            summary,
            diagnostics,
            inspection,
            assessment,
            before,
            after,
            changes,
            wasReserialized,
            wereInvalidatedSignaturesRemoved);

    private sealed record ValidatedProvenanceRequest(
        string Id,
        OfficeProvenanceWorkflowOperation Operation,
        string InputPath,
        string? OutputPath,
        ProvenanceOwner Owner,
        OfficeWorkflowConflictPolicy ConflictPolicy,
        SortedSet<string>? BatchBlockedOutputIdentities,
        string? BatchOwnReservedOutputIdentity,
        OfficeWorkflowLimits Limits,
        OfficeProvenanceOptions Inspection,
        OfficeProvenanceAssessmentOptions Assessment,
        OfficeProvenanceRemovalOptions Removal,
        OfficeProvenanceOptions RemovalInputInspection,
        OfficeProvenanceOptions RemovalOutputInspection);
}
