using OfficeIMO.Tool.Commands.Reader;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    // Runs one catalog route through the shared Workflows runner, the same registration the CLI and Workflows hosts use.
    private async Task<AgentConvertResult> ConvertRouteAsync(string path, string outputPath, string format, bool overwrite,
        long maximumInputBytes, long maximumOutputBytes, CancellationToken token) {
        string input = _pathPolicy.ResolveInput(path);
        if (Directory.Exists(input)) throw new AgentUsageException("Document conversion requires one local file; directory packages are limited to iWork sources.");
        string destination = _pathPolicy.ResolveOutput(outputPath);
        ReaderToolPathSafety.EnsureDistinctFile(input, destination);
        string target = "." + format;
        if (!Path.GetExtension(destination).Equals(target, StringComparison.OrdinalIgnoreCase))
            throw new AgentUsageException("The output path extension must match the requested conversion format.");
        var runner = new OfficeWorkflowRunner();
        string extension = Path.GetExtension(input).ToLowerInvariant();
        OfficeWorkflowRoute route = (extension.Length == 0 ? null : OfficeWorkflowCatalog.Find(extension, target, runner.ConversionRoutes))
            ?? throw new AgentUsageException("No conversion route converts '" + extension + "' to '" + target
                + "'. Use officeimo_capabilities with operation 'convert' to list the supported routes.");
        var request = new OfficeWorkflowRequest {
            InputPath = input, OutputPath = destination, Operation = OfficeWorkflowOperation.Convert,
            ConversionRouteId = route.Id,
            ConflictPolicy = overwrite ? OfficePublicationPolicy.Replace : OfficePublicationPolicy.Fail,
            Limits = new OfficeOperationLimits { MaximumInputBytes = maximumInputBytes, MaximumOutputBytes = maximumOutputBytes },
            PublicationGuard = new AgentConversionPublicationGuard(_pathPolicy, input)
        };
        OfficeWorkflowResult workflow = await runner.RunAsync(request, cancellationToken: token).ConfigureAwait(false);
        return new AgentConvertResult {
            SourcePath = input, OutputPath = workflow.OutputPath ?? string.Empty, Format = format,
            LengthBytes = workflow.OutputBytes, Sha256 = workflow.OutputSha256 ?? string.Empty,
            DiagnosticCount = workflow.Diagnostics.Count, Workflow = workflow
        };
    }
}
