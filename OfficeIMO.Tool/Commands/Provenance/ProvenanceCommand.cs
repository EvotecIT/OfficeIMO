using System.Text.Json;
using OfficeIMO.Provenance.C2pa;
using OfficeIMO;
using OfficeIMO.Provenance;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Provenance;

internal static class ProvenanceCommand {
    internal const string Usage = """
OfficeIMO.Tool - provenance workflows

Usage:
  officeimo provenance doctor --c2patool <trusted-executable> [--format json|text]
  officeimo provenance capabilities [--format json|text]
  officeimo provenance inspect <input> [--no-embedded] [--max-input-bytes <bytes>] [--format json|text]
  officeimo provenance assess <input> [--no-embedded] [--no-text-integrity]
             [--max-input-bytes <bytes>] [--format json|text]
             [--c2patool <trusted-executable>] [--trust-anchors <pem>] [--allowed-list <pem>]
             [--verification-timeout-seconds <1-300>]
  officeimo provenance remove <input> [--output <path>] [--force]
             [--keep-c2pa] [--keep-external-c2pa] [--keep-ai-source]
             [--remove-invalidated-signatures] [--no-embedded]
             [--max-input-bytes <bytes>] [--max-output-bytes <bytes>] [--format json|text]
  officeimo provenance audit <file-or-directory>... [--include <wildcard>] [--exclude <wildcard>]
             [--no-recursive] [--max-items <1-10000>] [--format json|text|ndjson|sarif]
  officeimo provenance check <file-or-directory>... [--fail-on dangerous-text|carriers|any] [audit options]
  officeimo provenance batch inspect|assess <input>... [--max-items <1-10000>] [options]
  officeimo provenance batch remove <input>... --output-directory <path>
             [--max-items <1-10000>] [options]

Provider options apply to assess and batch assess. Verification stays offline; c2patool is host-supplied.
doctor executes --version only; availability does not establish signer trust.
audit and check never modify files. Directory discovery skips symbolic links, .git, bin, obj and node_modules.
check defaults to dangerous Unicode findings; exit 1 means the selected evidence policy found findings.
Missing inputs, incomplete discovery and failed assessments return their existing nonzero error codes.
JSON is the default output and carries a versioned schema identifier.
Removal preserves the input format, refuses existing output unless --force is supplied,
and blocks invalidating package signatures unless --remove-invalidated-signatures is explicit.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default,
        IOfficeProvenanceWorkflowRunner? runner = null) {
        try {
            ProvenanceArguments parsed = ProvenanceArguments.Parse(args);
            if (parsed.Command == ProvenanceCommandKind.Help) {
                await standardOutput.WriteLineAsync(Usage).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            if (parsed.Command == ProvenanceCommandKind.Capabilities) {
                await ProvenanceOutput.WriteCapabilitiesAsync(standardOutput, parsed.Format).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }

            if (parsed.Command == ProvenanceCommandKind.Doctor) {
                C2paToolAvailability readiness = new C2paToolProvenanceVerifier(parsed.C2paToolPath!).CheckAvailability(
                    TimeSpan.FromSeconds(parsed.VerificationTimeoutSeconds), cancellationToken);
                await standardOutput.WriteLineAsync(parsed.Format == ProvenanceOutputFormat.Json
                    ? JsonSerializer.Serialize(new ProvenanceDoctorDto("officeimo.provenance.doctor.v1", readiness.Available, readiness.ExecutablePath, readiness.Version, readiness.Diagnostic), ProvenanceJsonContext.Default.ProvenanceDoctorDto)
                    : $"{(readiness.Available ? "Available" : "Unavailable")}: {readiness.Version ?? "unknown version"}. {readiness.Diagnostic}").ConfigureAwait(false);
                return readiness.Available ? (int)OfficeImoToolExitCode.Success : (int)OfficeImoToolExitCode.OperationFailed;
            }
            IOfficeProvenanceWorkflowRunner activeRunner = runner ?? new OfficeWorkflowRunner(
                provenanceVerifier: parsed.C2paToolPath == null ? null : new C2paToolProvenanceVerifier(parsed.C2paToolPath));
            if (parsed.Command is ProvenanceCommandKind.Audit or ProvenanceCommandKind.Check) {
                var audit = new OfficeProvenanceAuditRequest { Inputs = parsed.Inputs, Recursive = parsed.Recursive,
                    Include = parsed.Include, Exclude = parsed.Exclude, MaximumItems = parsed.MaximumItems,
                    MaximumInputBytes = parsed.MaximumInputBytes, InspectTextIntegrity = parsed.InspectTextIntegrity,
                    ProcessEmbeddedAssets = parsed.ProcessEmbeddedAssets };
                IReadOnlyList<OfficeProvenanceWorkflowResult> reports = await OfficeProvenanceAudit.RunAsync(audit, activeRunner, cancellationToken).ConfigureAwait(false);
                if (parsed.Format == ProvenanceOutputFormat.Ndjson) {
                    foreach (var report in reports) await standardOutput.WriteLineAsync(OfficeProvenanceReportSerializer.Serialize(report)).ConfigureAwait(false);
                } else if (parsed.Format == ProvenanceOutputFormat.Sarif) {
                    await standardOutput.WriteLineAsync(OfficeProvenanceSarif.Serialize(reports)).ConfigureAwait(false);
                } else await ProvenanceOutput.WriteBatchAsync(standardOutput, reports, parsed.Format).ConfigureAwait(false);
                int execution = MapBatch(reports);
                if (execution != 0) return execution;
                return parsed.Command == ProvenanceCommandKind.Check && reports.Any(report =>
                    OfficeProvenanceAudit.HasFindings(report, parsed.FailOnCarriers, parsed.FailOnDangerousText))
                    ? (int)OfficeImoToolExitCode.ValidationFailed : (int)OfficeImoToolExitCode.Success;
            }
            if (parsed.Command == ProvenanceCommandKind.Batch) {
                if (cancellationToken.IsCancellationRequested) {
                    OfficeProvenanceWorkflowResult cancelled = await activeRunner.RunProvenanceAsync(
                        CreateBatchRequest(parsed, parsed.Inputs[0]),
                        cancellationToken: cancellationToken).ConfigureAwait(false);
                    IReadOnlyList<OfficeProvenanceWorkflowResult> cancelledResults = new[] { cancelled };
                    await ProvenanceOutput.WriteBatchAsync(standardOutput, cancelledResults, parsed.Format).ConfigureAwait(false);
                    return MapBatch(cancelledResults);
                }
                IReadOnlyList<OfficeProvenanceWorkflowResult> results = await activeRunner.RunProvenanceBatchAsync(
                    CreateBatchRequests(parsed),
                    new OfficeProvenanceWorkflowBatchOptions { MaximumRequests = parsed.MaximumItems },
                    cancellationToken: cancellationToken).ConfigureAwait(false);
                await ProvenanceOutput.WriteBatchAsync(standardOutput, results, parsed.Format).ConfigureAwait(false);
                return MapBatch(results);
            }

            OfficeProvenanceWorkflowResult result = await activeRunner.RunProvenanceAsync(
                CreateRequest(parsed, parsed.Inputs[0], parsed.OutputPath),
                cancellationToken: cancellationToken).ConfigureAwait(false);
            await ProvenanceOutput.WriteResultAsync(standardOutput, result, parsed.Format).ConfigureAwait(false);
            return MapBatch(new[] { result });
        } catch (ProvenanceUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (ArgumentException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Provenance workflow cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            await standardError.WriteLineAsync(
                "Provenance workflow failed: " + exception.GetType().Name + ": " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static IEnumerable<OfficeProvenanceWorkflowRequest> CreateBatchRequests(ProvenanceArguments parsed) {
        foreach (string input in parsed.Inputs) {
            yield return CreateBatchRequest(parsed, input);
        }
    }

    private static OfficeProvenanceWorkflowRequest CreateBatchRequest(ProvenanceArguments parsed, string input) {
        string fullInput = Path.GetFullPath(input);
        string? outputDirectory = parsed.OutputDirectory is null ? null : Path.GetFullPath(parsed.OutputDirectory);
        string? output = outputDirectory is null
            ? null
            : Path.Combine(
                outputDirectory,
                Path.GetFileNameWithoutExtension(fullInput) + ".provenance-cleaned" + Path.GetExtension(fullInput));
        return CreateRequest(parsed, fullInput, output, parsed.BatchOperation);
    }

    internal static OfficeProvenanceWorkflowRequest CreateRequest(
        ProvenanceArguments parsed,
        string input,
        string? output,
        OfficeProvenanceWorkflowOperation? operation = null) {
        var request = new OfficeProvenanceWorkflowRequest {
            Operation = operation ?? parsed.Command switch {
                ProvenanceCommandKind.Inspect => OfficeProvenanceWorkflowOperation.Inspect,
                ProvenanceCommandKind.Assess => OfficeProvenanceWorkflowOperation.Assess,
                ProvenanceCommandKind.Remove => OfficeProvenanceWorkflowOperation.Remove,
                _ => throw new InvalidOperationException("The command does not map to one provenance operation.")
            },
            InputPath = Path.GetFullPath(input),
            OutputPath = output is null ? null : Path.GetFullPath(output),
            ConflictPolicy = parsed.Force ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail,
            Limits = new OfficeWorkflowLimits {
                MaximumInputBytes = parsed.MaximumInputBytes,
                MaximumOutputBytes = parsed.MaximumOutputBytes
            }
        };
        long parserInputBytes = Math.Min(parsed.MaximumInputBytes, int.MaxValue);
        long parserOutputBytes = Math.Min(parsed.MaximumOutputBytes, int.MaxValue);
        request.Inspection.MaxAssetBytes = parserInputBytes;
        request.Inspection.ProcessEmbeddedAssets = parsed.ProcessEmbeddedAssets;
        request.Assessment.Structural.MaxAssetBytes = parserInputBytes;
        request.Assessment.Structural.ProcessEmbeddedAssets = parsed.ProcessEmbeddedAssets;
        request.Assessment.TextIntegrity.MaxEncodedBytes = parserInputBytes;
        request.Assessment.InspectTextIntegrity = parsed.InspectTextIntegrity;
        request.Assessment.Verification.Timeout = TimeSpan.FromSeconds(parsed.VerificationTimeoutSeconds);
        request.Assessment.Verification.TrustAnchorsPath = parsed.TrustAnchorsPath;
        request.Assessment.Verification.AllowedListPath = parsed.AllowedListPath;
        request.Removal.Limits.MaxAssetBytes = parserInputBytes;
        request.Removal.MaxOutputBytes = parserOutputBytes;
        request.Removal.RemoveC2paManifests = parsed.RemoveC2paManifests;
        request.Removal.RemoveExternalC2paReferences = parsed.RemoveExternalC2paReferences;
        request.Removal.RemoveAiSourceMetadata = parsed.RemoveAiSourceMetadata;
        request.Removal.ProcessEmbeddedAssets = parsed.ProcessEmbeddedAssets;
        request.Removal.SignatureMutationPolicy = parsed.RemoveInvalidatedSignatures
            ? OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            : OfficeSignatureMutationPolicy.BlockSave;
        return request;
    }

    private static int MapBatch(IReadOnlyList<OfficeProvenanceWorkflowResult> results) {
        if (results.Any(result => result.Status == OfficeWorkflowStatus.Cancelled)) {
            return (int)OfficeImoToolExitCode.Cancelled;
        }
        if (results.Any(result => result.Assessment?.VerificationStatus == OfficeProvenanceCheckStatus.Failed ||
            result.Assessment?.ProviderSignalsStatus == OfficeProvenanceCheckStatus.Failed))
            return (int)OfficeImoToolExitCode.OperationFailed;
        OfficeProvenanceWorkflowResult? failed = results.FirstOrDefault(result => !result.Succeeded);
        return failed is null
            ? (int)OfficeImoToolExitCode.Success
            : MapStatus(failed.Status, failed.FailureKind);
    }

    private static int MapStatus(OfficeWorkflowStatus status, OfficeWorkflowFailureKind failureKind) => status switch {
        OfficeWorkflowStatus.Completed => (int)OfficeImoToolExitCode.Success,
        OfficeWorkflowStatus.Cancelled => (int)OfficeImoToolExitCode.Cancelled,
        OfficeWorkflowStatus.Failed => failureKind switch {
            OfficeWorkflowFailureKind.ValidationFailed => (int)OfficeImoToolExitCode.Usage,
            OfficeWorkflowFailureKind.InputNotFound => (int)OfficeImoToolExitCode.InputNotFound,
            OfficeWorkflowFailureKind.UnsupportedInput => (int)OfficeImoToolExitCode.UnsupportedInput,
            OfficeWorkflowFailureKind.OutputFailed => (int)OfficeImoToolExitCode.OutputFailed,
            _ => (int)OfficeImoToolExitCode.OperationFailed
        },
        _ => (int)OfficeImoToolExitCode.OperationFailed
    };
}
