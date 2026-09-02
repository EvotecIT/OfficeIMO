using System.Diagnostics;
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
        long inputBytes = 0;
        WorkflowFailureStage failureStage = WorkflowFailureStage.Validation;

        try {
            validated = ValidateProvenanceRequest(request);
            string ownerPackage = GetPackage(validated.Owner);
            Report(progress, validated.Id, "validate", "Validating provenance input and limits", 0.05D);
            cancellationToken.ThrowIfCancellationRequested();
            failureStage = WorkflowFailureStage.Input;
            inputBytes = new FileInfo(validated.InputPath).Length;
            EnforceInputLimit(validated.InputPath, inputBytes, validated.Limits);

            Report(progress, validated.Id, "inspect", "Inspecting through " + ownerPackage, 0.2D);
            failureStage = WorkflowFailureStage.Operation;
            OfficeProvenanceOptions inspectionOptions = validated.Operation == OfficeProvenanceWorkflowOperation.Assess
                ? validated.Assessment.Structural
                : validated.Inspection;
            OfficeProvenanceReport structural = await Task.Run(
                () => OfficeProvenanceWorkflowAdapter.Inspect(validated.Owner, validated.InputPath, inspectionOptions),
                cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            ProvenanceOwner refinedOwner = Refine(validated.Owner, structural.Format);
            if (structural.Format == OfficeProvenanceAssetFormat.Unknown) {
                throw new NotSupportedException("The input is not a supported provenance asset.");
            }
            validated = validated with { Owner = refinedOwner };
            ownerPackage = GetPackage(refinedOwner);

            if (validated.Operation == OfficeProvenanceWorkflowOperation.Inspect) {
                Report(progress, validated.Id, "complete", "Structural provenance report is ready", 1D);
                return CreateProvenanceResult(
                    validated, OfficeWorkflowStatus.Completed, OfficeWorkflowFailureKind.None,
                    ownerPackage, null, inputBytes, 0, stopwatch.Elapsed,
                    DescribeInspection(structural), diagnostics, inspection: structural);
            }

            if (validated.Operation == OfficeProvenanceWorkflowOperation.Assess) {
                Report(progress, validated.Id, "assess", "Collecting optional verification and signal evidence", 0.55D);
                OfficeProvenanceAssessmentReport assessment = await Task.Run(
                    () => OfficeProvenanceAssessment.AssessFile(
                        validated.InputPath,
                        structural,
                        validated.Assessment,
                        _provenanceVerifier,
                        _provenanceSignalDetectors,
                        cancellationToken),
                    cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
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

            failureStage = WorkflowFailureStage.Output;
            string outputDirectory = Path.GetDirectoryName(validated.OutputPath!)!;
            Directory.CreateDirectory(outputDirectory);
            stagingPath = Path.Combine(
                outputDirectory,
                "." + Path.GetFileNameWithoutExtension(validated.OutputPath) + "." +
                Guid.NewGuid().ToString("N") + Path.GetExtension(validated.OutputPath));
            Report(progress, validated.Id, "remove", "Removing selected carriers through " + ownerPackage, 0.48D);
            failureStage = WorkflowFailureStage.Operation;
            OfficeProvenanceRemovalResult removal = await Task.Run(
                () => OfficeProvenanceWorkflowAdapter.Remove(refinedOwner, validated.InputPath, stagingPath, validated.Removal),
                cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();

            failureStage = WorkflowFailureStage.Output;
            long stagedBytes = new FileInfo(stagingPath).Length;
            if (stagedBytes > validated.Limits.MaximumOutputBytes) {
                throw new InvalidOperationException(
                    $"Generated artifact is {stagedBytes:N0} bytes, above the configured {validated.Limits.MaximumOutputBytes:N0}-byte limit.");
            }

            Report(progress, validated.Id, "validate-output", "Reopening the staged artifact through " + ownerPackage, 0.72D);
            OfficeProvenanceReport reopened = await Task.Run(
                () => OfficeProvenanceWorkflowAdapter.Inspect(refinedOwner, stagingPath, CreateInspectionOptions(validated.Removal, validated.Limits)),
                cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            EnsureEquivalent(removal.After, reopened);
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
            Report(progress, validated.Id, "publish", "Publishing the verified provenance artifact", 0.9D);
            string publishedPath = Publish(stagingPath, validated.OutputPath!, validated.ConflictPolicy, cancellationToken);
            stagingPath = null;
            long outputBytes = new FileInfo(publishedPath).Length;
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "AtomicPublication",
                "The verified artifact was staged in the destination directory and published with one filesystem move.",
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
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "Cancelled",
                "The provenance workflow was cancelled before publication; no staged artifact was retained.",
                OfficeWorkflowDiagnosticSeverity.Information,
                "cancel"));
            return new OfficeProvenanceWorkflowResult(
                validated?.Id ?? request.Id,
                validated?.Operation ?? request.Operation,
                OfficeWorkflowStatus.Cancelled,
                OfficeWorkflowFailureKind.None,
                GetPackage(validated?.Owner ?? ResolveByPath(request.InputPath)),
                null,
                inputBytes,
                0,
                stopwatch.Elapsed,
                "Cancelled",
                diagnostics);
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
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
                GetPackage(validated?.Owner ?? ResolveByPath(request.InputPath)),
                null,
                inputBytes,
                0,
                stopwatch.Elapsed,
                "Provenance workflow failed: " + exception.Message,
                diagnostics);
        } finally {
            if (stagingPath is not null) TryDelete(stagingPath);
        }
    }

    /// <summary>Runs a bounded batch sequentially so providers and parsers share one predictable resource envelope.</summary>
    public async Task<IReadOnlyList<OfficeProvenanceWorkflowResult>> RunProvenanceBatchAsync(
        IEnumerable<OfficeProvenanceWorkflowRequest> requests,
        OfficeProvenanceWorkflowBatchOptions? options = null,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(requests);
        OfficeProvenanceWorkflowBatchOptions validatedOptions = (options ?? new OfficeProvenanceWorkflowBatchOptions()).CloneAndValidate();
        OfficeProvenanceWorkflowRequest[] materialized = requests.Take(validatedOptions.MaximumRequests + 1).ToArray();
        if (materialized.Length > validatedOptions.MaximumRequests) {
            throw new ArgumentException(
                $"The provenance batch exceeds the configured limit of {validatedOptions.MaximumRequests:N0} requests.",
                nameof(requests));
        }
        EnsureDistinctBatchRemovalOutputs(materialized);
        if (materialized.Length == 0) cancellationToken.ThrowIfCancellationRequested();

        var results = new List<OfficeProvenanceWorkflowResult>(materialized.Length);
        for (int index = 0; index < materialized.Length; index++) {
            int batchIndex = index;
            var batchProgress = progress is null
                ? null
                : new InlineProgress<OfficeWorkflowProgress>(item => progress.Report(new OfficeWorkflowProgress(
                    item.RequestId,
                    item.Stage,
                    $"{batchIndex + 1} of {materialized.Length} · {item.Message}",
                    item.Fraction,
                    (batchIndex + item.Fraction) / Math.Max(1, materialized.Length))));
            OfficeProvenanceWorkflowResult result = await RunProvenanceAsync(
                materialized[index], batchProgress, cancellationToken).ConfigureAwait(false);
            results.Add(result);
            if (result.Status == OfficeWorkflowStatus.Cancelled) break;
            if (!validatedOptions.ContinueOnFailure && !result.Succeeded) break;
        }
        return results;
    }

    private static void EnsureDistinctBatchRemovalOutputs(IEnumerable<OfficeProvenanceWorkflowRequest> requests) {
        var outputIdentities = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeProvenanceWorkflowRequest request in requests) {
            if (request == null) throw new ArgumentException("Provenance batches cannot contain null requests.", nameof(requests));
            if (request.Operation != OfficeProvenanceWorkflowOperation.Remove) continue;

            string? outputPath = TryResolveBatchRemovalOutput(request);
            if (outputPath is null) continue;
            string? identity = TryNormalizeBatchRemovalOutput(outputPath);
            if (identity is null) continue;
            if (!outputIdentities.Add(identity)) {
                throw new ArgumentException(
                    $"Multiple provenance removal requests resolve to the same output path '{outputPath}'.",
                    nameof(requests));
            }
        }
    }

    private static string? TryNormalizeBatchRemovalOutput(string outputPath) {
        try {
            return OfficeWorkflowPathIdentity.Normalize(outputPath);
        } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
            // Defer path-access diagnostics to the item runner so batch failure behavior stays structured.
            return null;
        }
    }

    private static string? TryResolveBatchRemovalOutput(OfficeProvenanceWorkflowRequest request) {
        try {
            if (string.IsNullOrWhiteSpace(request.InputPath)) return null;
            string inputPath = Path.GetFullPath(request.InputPath);
            return string.IsNullOrWhiteSpace(request.OutputPath)
                ? Path.Combine(
                    Path.GetDirectoryName(inputPath)!,
                    Path.GetFileNameWithoutExtension(inputPath) + ".provenance-cleaned" + Path.GetExtension(inputPath))
                : Path.GetFullPath(request.OutputPath);
        } catch (Exception exception) when (exception is ArgumentException or NotSupportedException or IOException or UnauthorizedAccessException) {
            // RunProvenanceAsync owns request validation and converts invalid paths into per-item results.
            return null;
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
        OfficeProvenanceRemovalOptions removal = CloneRemovalOptions(
            request.Removal ?? throw new ArgumentException("Removal options cannot be null.", nameof(request)),
            limits);
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
            limits,
            inspection,
            assessment,
            removal);
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
        OfficeWorkflowLimits limits) {
        var clone = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            SignatureMutationPolicy = source.SignatureMutationPolicy,
            ProcessEmbeddedAssets = source.ProcessEmbeddedAssets,
            MaxEmbeddedAssets = source.MaxEmbeddedAssets
        };
        CopyInspectionOptions(source.Limits, clone.Limits, limits.MaximumInputBytes);
        return clone;
    }

    private static OfficeProvenanceOptions CreateInspectionOptions(
        OfficeProvenanceRemovalOptions removal,
        OfficeWorkflowLimits limits) {
        var options = new OfficeProvenanceOptions();
        CopyInspectionOptions(removal.Limits, options, limits.MaximumOutputBytes);
        options.ProcessEmbeddedAssets = removal.ProcessEmbeddedAssets && removal.Limits.ProcessEmbeddedAssets;
        options.MaxEmbeddedAssets = Math.Min(removal.MaxEmbeddedAssets, removal.Limits.MaxEmbeddedAssets);
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
        OfficeWorkflowLimits Limits,
        OfficeProvenanceOptions Inspection,
        OfficeProvenanceAssessmentOptions Assessment,
        OfficeProvenanceRemovalOptions Removal);
}
