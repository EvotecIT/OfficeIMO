using OfficeIMO.Tool.Agent;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

internal static class WorkflowCompareCommand {
    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken token) {
        var inputs = new List<string>();
        string? destination = null, expectedPages = null, actualPages = null, passwordName = null, actualPasswordName = null;
        bool overwrite = false;
        for (int i = 0; i < args.Length; i++) {
            string option = args[i];
            string Value() => ++i < args.Length && !args[i].StartsWith("--", StringComparison.Ordinal)
                ? args[i] : throw new WorkflowUsageException(option + " requires a value.");
            switch (option) {
                case "--output": destination = Value(); break;
                case "--expected-pages": expectedPages = Value(); break;
                case "--actual-pages": actualPages = Value(); break;
                case "--password-env": passwordName = Value(); break;
                case "--comparison-password-env": actualPasswordName = Value(); break;
                case "--force": overwrite = true; break;
                default:
                    if (option.StartsWith("-", StringComparison.Ordinal)) throw new WorkflowUsageException("Unknown comparison option " + option + ".");
                    inputs.Add(option); break;
            }
        }
        if (inputs.Count != 2 || string.IsNullOrWhiteSpace(destination))
            throw new WorkflowUsageException("compare requires two PDF inputs and --output <report.html>.");
        var expected = new PdfWorkflowSettings { Pages = expectedPages, PasswordEnvironmentVariable = passwordName };
        var actual = new PdfWorkflowSettings { Pages = actualPages, PasswordEnvironmentVariable = actualPasswordName };
        try { expected.Validate(); actual.Validate(); }
        catch (AgentUsageException exception) { throw new WorkflowUsageException(exception.Message); }
        OfficeWorkflowResult result;
        try {
            result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Compare,
                InputPath = Path.GetFullPath(inputs[0]), ComparisonPath = Path.GetFullPath(inputs[1]), OutputPath = Path.GetFullPath(destination),
                ComparisonExpectedPages = expected.Selector(), ComparisonActualPages = actual.Selector(),
                PdfPassword = expected.Password(), ComparisonPdfPassword = actual.Password(),
                Limits = new OfficeWorkflowLimits { MaximumInputBytes = expected.MaximumInputBytes, MaximumOutputBytes = 96L * 1024 * 1024 },
                ConflictPolicy = overwrite ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail
            }, cancellationToken: token).ConfigureAwait(false);
        } catch (AgentUsageException exception) { throw new WorkflowUsageException(exception.Message); }
        foreach (var diagnostic in result.Diagnostics) await error.WriteLineAsync(diagnostic.Code + ": " + diagnostic.Message).ConfigureAwait(false);
        if (!result.Succeeded) return WorkflowCommand.MapStatus(result.Status, result.FailureKind);
        await output.WriteLineAsync(result.OutputPath).ConfigureAwait(false);
        await output.WriteLineAsync(result.Summary).ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Success;
    }
}
