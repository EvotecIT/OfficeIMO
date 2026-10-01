using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

internal static class WorkflowArchiveCommand {
    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken token) {
        if (args.Length == 1 && args[0] is "--help" or "-h") {
            await output.WriteLineAsync(WorkflowCommand.Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        if (args.Length is not 2 and not 3 || args[0] != "--request" || (args.Length == 3 && args[2] != "--retry-failed"))
            throw new WorkflowUsageException("Use archive --request archive.json [--retry-failed].");
        await using var source = File.OpenRead(args[1]);
        if (source.Length > 64 * 1024) throw new WorkflowUsageException("Archive request exceeds 64 KiB.");
        var bytes = new byte[checked((int)source.Length)];
        await source.ReadExactlyAsync(bytes, token).ConfigureAwait(false);
        OfficePdfArchiveRequest request = OfficePdfArchiveSerializer.ParseRequest(bytes);
        if (args.Length == 3) request = request with { RetryFailed = true };
        var progress = new ArchiveProgress(TextWriter.Synchronized(error));
        OfficePdfArchiveResult result = await OfficePdfArchiveWorkflow.RunAsync(request, progress, token).ConfigureAwait(false);
        await output.WriteLineAsync(OfficePdfArchiveSerializer.SerializeResult(result)).ConfigureAwait(false);
        return (int)(result.Cancelled ? OfficeImoToolExitCode.Cancelled : result.Failed > 0 ? OfficeImoToolExitCode.OperationFailed : OfficeImoToolExitCode.Success);
    }

    private sealed class ArchiveProgress(TextWriter output) : IProgress<OfficePdfArchiveItemResult> {
        public void Report(OfficePdfArchiveItemResult item) {
            if (item.Status != OfficeWorkflowStatus.Completed) output.WriteLine(item.InputPath + ": " + item.Summary);
            foreach (OfficeWorkflowDiagnostic finding in item.Diagnostics.Where(finding => finding.Severity != OfficeWorkflowDiagnosticSeverity.Information))
                output.WriteLine(item.InputPath + ": " + finding.Code + ": " + finding.Message);
        }
    }
}
