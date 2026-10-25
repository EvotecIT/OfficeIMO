using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace OfficeIMO.Tool.Commands.Workflow;

internal static class WorkflowBatchCommand {
    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken token) {
        if (args.Length == 1 && args[0] is "--help" or "-h") {
            await output.WriteLineAsync(WorkflowCommand.Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        WorkflowBatchArguments parsed = Parse(args);
        var settings = parsed.Settings;
        if (parsed.SaveJobPath is { } jobPath) {
            string fullPath = Path.GetFullPath(jobPath);
            string sourceBase = Environment.CurrentDirectory;
            settings = settings with {
                InputDirectory = settings.InputDirectory == null ? null : Path.GetFullPath(settings.InputDirectory, sourceBase),
                InputPaths = settings.InputPaths?.Select(path => Path.GetFullPath(path, sourceBase)).ToArray(),
                OutputDirectory = Path.GetFullPath(settings.OutputDirectory, sourceBase),
                CheckpointDirectory = settings.CheckpointDirectory == null ? null : Path.GetFullPath(settings.CheckpointDirectory, sourceBase)
            };
            OfficeConversionJob job = OfficeConversionJob.FromRequest(settings, Path.GetDirectoryName(fullPath)!);
            job.Secrets.SourcePasswordEnvironmentVariable = parsed.SourcePasswordVariable;
            job.Secrets.PdfPasswordEnvironmentVariable = parsed.PdfPasswordVariable;
            token.ThrowIfCancellationRequested();
            OfficeConversionJobSerializer.SaveFile(job, fullPath);
            await output.WriteLineAsync(fullPath).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        settings.ConversionOptions.SourcePassword ??= ReadSecret(parsed.SourcePasswordVariable);
        settings.PdfPassword ??= ReadSecret(parsed.PdfPasswordVariable);
        var progress = new WorkflowBatchOutput(output, error, parsed.Format);
        var result = await IWorkWorkflow.CreateRunner().RunBatchAsync(settings, progress, token).ConfigureAwait(false);
        progress.Complete(result, parsed.AllowEmpty);
        return ExitCode(result, parsed.AllowEmpty);
    }

    internal static WorkflowBatchArguments Parse(string[] args) {
        var paths = new List<string>();
        var extensions = new List<string>();
        var settings = new OfficeConversionBatchRequest { OutputDirectory = "" };
        var text = new PdfPlainTextOptions();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        bool allowEmpty = false;
        string? format = null, saveJob = null, sourcePasswordVariable = null, pdfPasswordVariable = null;
        for (int index = 0; index < args.Length; index++) {
            string option = args[index];
            if (!option.StartsWith("--", StringComparison.Ordinal)) { paths.Add(option); continue; }
            if (option != "--source-extension" && !seen.Add(option)) throw new WorkflowUsageException("Duplicate batch option: " + option);
            if (option == "--retry-failed") { settings.RetryFailed = true; continue; }
            if (option == "--no-recursive") { settings.Recursive = false; continue; }
            if (option == "--allow-empty") { allowEmpty = true; continue; }
            if (option == "--allow-legacy-loss") { settings.ConversionOptions.LegacyDocLossPolicy = OfficeIMO.OfficeConversionLossPolicy.Allow; continue; }
            if (option == "--require-no-loss") { settings.ConversionOptions.RequireNoLoss = true; continue; }
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
                case "--source-password-env": sourcePasswordVariable = value; break;
                case "--pdf-password-env": pdfPasswordVariable = value; break;
                case "--format": format = value switch { "json" or "ndjson" => value, _ => throw new WorkflowUsageException("Choose batch format: json or ndjson.") }; break;
                case "--save-job": saveJob = value; break;
                case "--pages": settings.ConversionOptions.PageRanges = value; break;
                case "--source-extension": extensions.Add(value); break;
                case "--text-encoding": text.EncodingName = value; break;
                case "--tab-size": text.TabSize = Number(value, option, 32); break;
                case "--concurrency": settings.MaximumConcurrency = Number(value, option, 32); break;
                case "--maximum-files": settings.MaximumFiles = Number(value, option, int.MaxValue); break;
                case "--maximum-input-bytes": settings.MaximumInputBytes = Bytes(value, option); break;
                case "--maximum-output-bytes": settings.MaximumOutputBytes = Bytes(value, option); break;
                case "--maximum-xml-characters-in-part": settings.MaximumXmlCharactersInPart = Bytes(value, option); break;
                case "--diagram-layers":
                    settings.ConversionOptions.Draw = new OfficeIMO.OpenDocument.Odg.Pdf.OdgToPdfOptions { ForPrint = value switch {
                        "print" => true, "screen" => false, _ => throw new WorkflowUsageException("Choose diagram layers: screen or print.")
                    } };
                    break;
                case "--profile": settings.OutputProfile = EnumValue<OfficeWorkflowOutputProfile>(value, option); break;
                case "--conflict": settings.ConflictPolicy = EnumValue<OfficeWorkflowConflictPolicy>(value, option); break;
                default: throw new WorkflowUsageException("Unknown batch option: " + option);
            }
        }
        if (string.IsNullOrWhiteSpace(settings.OutputDirectory)) throw new WorkflowUsageException("Batch conversion requires --output.");
        settings.InputPaths = paths.Count == 0 ? null : paths.ToArray();
        settings.SourceExtensions = extensions.Count == 0 ? null : extensions.ToArray();
        settings.ConversionOptions.PlainText = text;
        if (sourcePasswordVariable != null && settings.ConversionOptions.SourcePassword != null)
            throw new WorkflowUsageException("Choose --source-password-env or --source-password, not both.");
        if (pdfPasswordVariable != null && settings.PdfPassword != null)
            throw new WorkflowUsageException("Choose --pdf-password-env or --pdf-password, not both.");
        return new(settings, allowEmpty, format, saveJob, sourcePasswordVariable, pdfPasswordVariable);
    }

    internal static int ExitCode(OfficeConversionBatchResult result, bool allowEmpty) =>
        (int)(result.Cancelled ? OfficeImoToolExitCode.Cancelled : result.FailureSummary != null || result.Failed > 0 || (!allowEmpty && result.Selected == 0)
            ? OfficeImoToolExitCode.OperationFailed : OfficeImoToolExitCode.Success);

    internal static string? ReadSecret(string? variableName) {
        if (variableName is null) return null;
        if (string.IsNullOrWhiteSpace(variableName) || variableName.Any(char.IsControl) || variableName.Contains('='))
            throw new WorkflowUsageException("Choose a valid password environment-variable name.");
        string? value = Environment.GetEnvironmentVariable(variableName);
        if (string.IsNullOrEmpty(value)) throw new WorkflowUsageException("Environment variable '" + variableName + "' is missing or empty.");
        return value;
    }

    private static int Number(string value, string option, int maximum) => int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int number)
        && number > 0 && number <= maximum ? number : throw new WorkflowUsageException("Invalid " + option);
    private static long Bytes(string value, string option) => long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long number)
        && number > 0 ? number : throw new WorkflowUsageException("Invalid " + option);
    private static T EnumValue<T>(string value, string option) where T : struct, Enum => Enum.TryParse<T>(value, true, out var parsed)
        && Enum.IsDefined(parsed) ? parsed : throw new WorkflowUsageException("Invalid " + option);

}

internal sealed record WorkflowBatchArguments(OfficeConversionBatchRequest Settings, bool AllowEmpty, string? Format,
    string? SaveJobPath, string? SourcePasswordVariable, string? PdfPasswordVariable);
