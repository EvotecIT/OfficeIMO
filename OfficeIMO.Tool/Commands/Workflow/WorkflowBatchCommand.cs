using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

internal static class WorkflowBatchCommand {
    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken token) {
        if (args.Length == 1 && args[0] is "--help" or "-h") {
            await output.WriteLineAsync(WorkflowCommand.Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        var paths = new List<string>();
        var extensions = new List<string>();
        var settings = new OfficeConversionBatchRequest { OutputDirectory = "" };
        var text = new PdfPlainTextOptions();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < args.Length; index++) {
            string option = args[index];
            if (!option.StartsWith("--", StringComparison.Ordinal)) { paths.Add(option); continue; }
            if (option != "--source-extension" && !seen.Add(option)) throw new WorkflowUsageException("Duplicate batch option: " + option);
            if (option == "--retry-failed") { settings.RetryFailed = true; continue; }
            if (option == "--no-recursive") { settings.Recursive = false; continue; }
            if (option == "--allow-legacy-loss") { settings.ConversionOptions.LegacyDocLossPolicy = OfficeIMO.OfficeConversionLossPolicy.Allow; continue; }
            if (++index >= args.Length) throw new WorkflowUsageException("Missing value for " + option);
            string value = args[index];
            switch (option) {
                case "--input-directory": settings.InputDirectory = value; break;
                case "--output": settings.OutputDirectory = value; break;
                case "--checkpoint": settings.CheckpointDirectory = value; break;
                case "--target": settings.TargetExtension = value; break;
                case "--route": settings.ConversionRouteId = value; break;
                case "--source-password": settings.ConversionOptions.SourcePassword = value; break;
                case "--pdf-password": settings.PdfPassword = value; break;
                case "--pages": settings.ConversionOptions.PageRanges = value; break;
                case "--source-extension": extensions.Add(value); break;
                case "--text-encoding": text.EncodingName = value; break;
                case "--tab-size": text.TabSize = Number(value, option, 32); break;
                case "--concurrency": settings.MaximumConcurrency = Number(value, option, 32); break;
                case "--maximum-files": settings.MaximumFiles = Number(value, option, int.MaxValue); break;
                case "--maximum-input-bytes": settings.MaximumInputBytes = Bytes(value, option); break;
                case "--maximum-output-bytes": settings.MaximumOutputBytes = Bytes(value, option); break;
                case "--profile": settings.OutputProfile = EnumValue<OfficeWorkflowOutputProfile>(value, option); break;
                case "--conflict": settings.ConflictPolicy = EnumValue<OfficeWorkflowConflictPolicy>(value, option); break;
                default: throw new WorkflowUsageException("Unknown batch option: " + option);
            }
        }
        if (string.IsNullOrWhiteSpace(settings.OutputDirectory)) throw new WorkflowUsageException("Batch conversion requires --output.");
        settings.InputPaths = paths.Count == 0 ? null : paths.ToArray();
        settings.SourceExtensions = extensions.Count == 0 ? null : extensions.ToArray();
        settings.ConversionOptions.PlainText = text;
        var progress = new BatchProgress(TextWriter.Synchronized(error));
        var result = await new OfficeWorkflowRunner().RunBatchAsync(settings, progress, token).ConfigureAwait(false);
        await output.WriteLineAsync(OfficeConversionBatchSerializer.SerializeResult(result)).ConfigureAwait(false);
        return (int)(result.Cancelled ? OfficeImoToolExitCode.Cancelled : result.Failed > 0 ? OfficeImoToolExitCode.OperationFailed : OfficeImoToolExitCode.Success);
    }

    private static int Number(string value, string option, int maximum) => int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int number)
        && number > 0 && number <= maximum ? number : throw new WorkflowUsageException("Invalid " + option);
    private static long Bytes(string value, string option) => long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long number)
        && number > 0 ? number : throw new WorkflowUsageException("Invalid " + option);
    private static T EnumValue<T>(string value, string option) where T : struct, Enum => Enum.TryParse<T>(value, true, out var parsed)
        && Enum.IsDefined(parsed) ? parsed : throw new WorkflowUsageException("Invalid " + option);

    private sealed class BatchProgress(TextWriter output) : IProgress<OfficeConversionBatchItemResult> {
        public void Report(OfficeConversionBatchItemResult item) {
            if (item.Skipped || item.Status != OfficeWorkflowStatus.Completed) output.WriteLine(item.InputPath + ": " + item.Summary);
            foreach (var finding in item.Diagnostics.Where(finding => finding.Severity != OfficeWorkflowDiagnosticSeverity.Information))
                output.WriteLine(item.InputPath + ": " + finding.Code + ": " + finding.Message);
        }
    }
}
