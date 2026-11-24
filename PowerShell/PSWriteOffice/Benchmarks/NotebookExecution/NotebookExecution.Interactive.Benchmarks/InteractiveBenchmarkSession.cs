using Microsoft.DotNet.Interactive.Commands;
using Microsoft.DotNet.Interactive.Events;
using Microsoft.DotNet.Interactive.PowerShell;

namespace NotebookExecution.Interactive.Benchmarks;

public sealed class InteractiveBenchmarkSession : IDisposable {
    private readonly string? _modulePath = Environment.GetEnvironmentVariable("PSModulePath");
    private readonly string? _distributionChannel = Environment.GetEnvironmentVariable("POWERSHELL_DISTRIBUTION_CHANNEL");
    private readonly PowerShellKernel _kernel;

    public InteractiveBenchmarkSession() {
        try {
            _kernel = new PowerShellKernel();
        } catch {
            RestoreEnvironment();
            throw;
        }
    }

    public async Task<string[]> ExecuteAsync(string[] sources) {
        var output = new string[sources.Length];
        for (int index = 0; index < sources.Length; index++) {
            var result = await _kernel.SendAsync(new SubmitCode(sources[index])).ConfigureAwait(false);
            var failure = result.Events.OfType<CommandFailed>().FirstOrDefault();
            if (failure != null)
                throw new InvalidOperationException(failure.Message);
            if (!result.Events.OfType<CommandSucceeded>().Any())
                throw new InvalidOperationException($"Cell {index} did not report a successful submission.");
            output[index] = string.Concat(result.Events.OfType<StandardOutputValueProduced>()
                .SelectMany(static item => item.FormattedValues)
                .Where(static value => value.MimeType == "text/plain")
                .Select(static value => value.Value));
        }
        return output;
    }

    public void Dispose() {
        try {
            _kernel.Dispose();
        } finally {
            RestoreEnvironment();
        }
    }

    private void RestoreEnvironment() {
        // The peer prepends its module path to process variables used by other engines.
        Environment.SetEnvironmentVariable("PSModulePath", _modulePath);
        Environment.SetEnvironmentVariable("POWERSHELL_DISTRIBUTION_CHANNEL", _distributionChannel);
    }
}
