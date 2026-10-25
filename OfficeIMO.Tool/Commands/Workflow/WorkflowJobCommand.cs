using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace OfficeIMO.Tool.Commands.Workflow;

/// <summary>Thin saved-job adapter over the shared planner and batch executor.</summary>
internal static class WorkflowJobCommand {
    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken token) {
        if (args.Length == 0 || args[0] is "--help" or "-h") {
            await output.WriteLineAsync(WorkflowCommand.Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        if (args[0] is not "plan" and not "run" and not "report") throw new WorkflowUsageException("Choose job plan, job run or job report.");
        string command = args[0];
        string? path = null, format = null, reportPath = null;
        bool retryFailed = false, allowEmpty = false;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 1; index < args.Length; index++) {
            string option = args[index];
            if (!option.StartsWith("--", StringComparison.Ordinal)) {
                if (path != null) throw new WorkflowUsageException("Specify one saved job file.");
                path = option;
                continue;
            }
            if (!seen.Add(option)) throw new WorkflowUsageException("Duplicate job option: " + option);
            if (command != "run") throw new WorkflowUsageException("Job plan and report inspection accept only their file path.");
            switch (option) {
                case "--retry-failed": retryFailed = true; break;
                case "--allow-empty": allowEmpty = true; break;
                case "--report":
                    if (++index >= args.Length || args[index].StartsWith("--", StringComparison.Ordinal))
                        throw new WorkflowUsageException("Specify a new run-report path.");
                    reportPath = args[index];
                    break;
                case "--format":
                    if (++index >= args.Length || args[index] is not "json" and not "ndjson")
                        throw new WorkflowUsageException("Choose job result format: json or ndjson.");
                    format = args[index];
                    break;
                default: throw new WorkflowUsageException("Unknown job option: " + option);
            }
        }
        if (string.IsNullOrWhiteSpace(path)) throw new WorkflowUsageException("Specify a saved job file.");
        string fullPath = Path.GetFullPath(path);
        if (command == "report") {
            var report = OfficeConversionJobRunReportSerializer.ReadFile(fullPath);
            await output.WriteLineAsync(OfficeConversionJobRunReportSerializer.Serialize(report)).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        string baseDirectory = Path.GetDirectoryName(fullPath)!;
        OfficeConversionJob job = OfficeConversionJobSerializer.ReadFile(fullPath);
        var runner = IWorkWorkflow.CreateRunner();
        if (command == "plan") {
            OfficeConversionJobPlan plan = await OfficeConversionJobRunner.PlanAsync(job, baseDirectory, runner,
                cancellationToken: token).ConfigureAwait(false);
            await output.WriteLineAsync(JsonSerializer.Serialize(plan, WorkflowJobJsonContext.Default.OfficeConversionJobPlan)).ConfigureAwait(false);
            return plan.CanRun ? (int)OfficeImoToolExitCode.Success : (int)OfficeImoToolExitCode.OperationFailed;
        }
        if (retryFailed) job.RetryFailed = true;
        var progress = new WorkflowBatchOutput(output, error, format ?? "json");
        var runtime = new OfficeConversionJobRuntimeOptions { ResolveEnvironmentVariable = Environment.GetEnvironmentVariable };
        OfficeConversionJobRunReport? recorded = null;
        OfficeConversionBatchResult result;
        if (reportPath == null) {
            result = await OfficeConversionJobRunner.RunAsync(job, baseDirectory, runner, runtime, progress, token).ConfigureAwait(false);
        } else {
            recorded = await OfficeConversionJobRunner.RunRecordedAsync(job, fullPath, Path.GetFullPath(reportPath), runner,
                runtime, progress, token, allowEmpty: allowEmpty).ConfigureAwait(false);
            result = recorded.Batch?.Summary ?? throw new InvalidOperationException("The run report has no conversion result.");
        }
        progress.Complete(result, allowEmpty, recorded?.RecordingError);
        return recorded?.RecordingError == null ? WorkflowBatchCommand.ExitCode(result, allowEmpty) : (int)OfficeImoToolExitCode.OperationFailed;
    }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, UseStringEnumConverter = true,
    GenerationMode = JsonSourceGenerationMode.Serialization)]
[JsonSerializable(typeof(OfficeConversionJobPlan))]
internal partial class WorkflowJobJsonContext : JsonSerializerContext { }
