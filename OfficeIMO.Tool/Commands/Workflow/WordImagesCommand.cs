using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

/// <summary>Thin command-line adapter for the shared Word image file workflow.</summary>
internal static class WordImagesCommand {
    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error,
        CancellationToken token = default, IOfficeWorkflowRunner? runner = null) {
        token.ThrowIfCancellationRequested();
        if (args.Any(argument => argument is "--help" or "-h")) {
            await output.WriteLineAsync(WorkflowCommand.Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        var inputs = new List<string>();
        var options = new WordImageOptimizationOptions();
        string? destination = null, directory = null, format = null;
        bool analyze = false, force = false;
        for (int i = 0; i < args.Length; i++) {
            string argument = args[i];
            string Value() => ++i < args.Length && !args[i].StartsWith("--", StringComparison.Ordinal)
                ? args[i] : throw new WorkflowUsageException(argument + " requires a value.");
            switch (argument.ToLowerInvariant()) {
                case "--analyze": analyze = true; break;
                case "--force": force = true; break;
                case "--output": destination = Value(); break;
                case "--output-directory": directory = Value(); break;
                case "--format": format = Value().ToLowerInvariant(); break;
                case "--mode": options.Mode = Value().ToLowerInvariant() switch {
                    "downsample" => OfficeImageOptimizationMode.Downsample,
                    "recompress" => OfficeImageOptimizationMode.Recompress,
                    "both" => OfficeImageOptimizationMode.DownsampleAndRecompress,
                    _ => throw new WorkflowUsageException("Choose downsample, recompress, or both.")
                }; break;
                case "--dpi": options.TargetDpi = Number(Value(), argument); break;
                case "--quality": options.JpegQuality = Number(Value(), argument); break;
                default:
                    if (argument.StartsWith("--", StringComparison.Ordinal)) throw new WorkflowUsageException("Unknown image optimization option " + argument + ".");
                    inputs.Add(Path.GetFullPath(argument)); break;
            }
        }
        if (inputs.Count == 0 || inputs.Count > OfficeWorkflowRunner.MaximumBatchRequestCount)
            throw new WorkflowUsageException("Supply between 1 and 250 Word input files.");
        if (destination != null && (inputs.Count != 1 || directory != null))
            throw new WorkflowUsageException("--output requires one input and cannot be combined with --output-directory.");
        if (format is not null && format is not ("docx" or "doc" or "pdf"))
            throw new WorkflowUsageException("Choose docx, doc, or pdf output.");
        if (analyze && (destination != null || directory != null || format != null || force))
            throw new WorkflowUsageException("Analysis does not accept output or overwrite options.");
        if (!analyze && destination == null && directory == null)
            throw new WorkflowUsageException("Choose --output or --output-directory for the optimized copy.");
        if (destination != null && format != null)
            throw new WorkflowUsageException("The --output file extension selects its format; omit --format.");
        try { options = options.Clone(); }
        catch (ArgumentException exception) { throw new WorkflowUsageException(exception.Message); }
        var requests = inputs.Select(input => new OfficeWorkflowRequest {
            Id = input,
            Operation = analyze ? OfficeWorkflowOperation.AnalyzeWordImages : OfficeWorkflowOperation.OptimizeWordImages,
            InputPath = input, WordImageOptimization = options,
            OutputPath = analyze ? null : destination != null ? Path.GetFullPath(destination)
                : Path.Combine(Path.GetFullPath(directory!), Path.GetFileNameWithoutExtension(input) + ".optimized." + (format ?? Path.GetExtension(input).TrimStart('.'))),
            ConflictPolicy = force ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail
        }).ToArray();
        if (!analyze && requests.Select(request => request.OutputPath!).Distinct(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal).Count() != requests.Length)
            throw new WorkflowUsageException("The batch inputs produce duplicate output filenames. Choose separate destinations or run them individually.");
        var results = await (runner ?? new OfficeWorkflowRunner()).RunBatchAsync(requests, cancellationToken: token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        foreach (var result in results) {
            await output.WriteLineAsync((result.OutputPath ?? result.RequestId) + ": " + result.Summary).ConfigureAwait(false);
            foreach (var diagnostic in result.Diagnostics)
                await (result.Succeeded ? output : error).WriteLineAsync(diagnostic.Code + ": " + diagnostic.Message).ConfigureAwait(false);
        }
        if (results.Any(result => result.Status == OfficeWorkflowStatus.Cancelled)) return (int)OfficeImoToolExitCode.Cancelled;
        var failed = results.FirstOrDefault(result => !result.Succeeded);
        return failed == null ? (int)OfficeImoToolExitCode.Success : WorkflowCommand.MapStatus(failed.Status, failed.FailureKind);
    }

    private static int Number(string text, string argument) => int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)
        ? value : throw new WorkflowUsageException("Supply an integer for " + argument + ".");
}
