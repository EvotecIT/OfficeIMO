using OfficeIMO;
using OfficeIMO.Provenance;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Provenance;

internal static class ProvenanceCommand {
    internal const string Usage = """
OfficeIMO.Tool - provenance workflows

Usage:
  officeimo provenance capabilities [--format json|text]
  officeimo provenance inspect <input> [--no-embedded] [--max-input-bytes <bytes>] [--format json|text]
  officeimo provenance assess <input> [--no-embedded] [--no-text-integrity]
             [--max-input-bytes <bytes>] [--format json|text]
  officeimo provenance remove <input> [--output <path>] [--force]
             [--keep-c2pa] [--keep-external-c2pa] [--keep-ai-source]
             [--remove-invalidated-signatures] [--no-embedded]
             [--max-input-bytes <bytes>] [--max-output-bytes <bytes>] [--format json|text]
  officeimo provenance batch inspect|assess <input>... [--max-items <1-10000>] [options]
  officeimo provenance batch remove <input>... --output-directory <path>
             [--max-items <1-10000>] [options]

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

            IOfficeProvenanceWorkflowRunner activeRunner = runner ?? new OfficeWorkflowRunner();
            if (parsed.Command == ProvenanceCommandKind.Batch) {
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
            return MapStatus(result.Status, result.FailureKind);
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
        string? outputDirectory = parsed.OutputDirectory is null ? null : Path.GetFullPath(parsed.OutputDirectory);
        foreach (string input in parsed.Inputs) {
            string fullInput = Path.GetFullPath(input);
            string? output = outputDirectory is null
                ? null
                : Path.Combine(
                    outputDirectory,
                    Path.GetFileNameWithoutExtension(fullInput) + ".provenance-cleaned" + Path.GetExtension(fullInput));
            yield return CreateRequest(parsed, fullInput, output, parsed.BatchOperation);
        }
    }

    private static OfficeProvenanceWorkflowRequest CreateRequest(
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
        request.Inspection.ProcessEmbeddedAssets = parsed.ProcessEmbeddedAssets;
        request.Assessment.Structural.ProcessEmbeddedAssets = parsed.ProcessEmbeddedAssets;
        request.Assessment.InspectTextIntegrity = parsed.InspectTextIntegrity;
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
