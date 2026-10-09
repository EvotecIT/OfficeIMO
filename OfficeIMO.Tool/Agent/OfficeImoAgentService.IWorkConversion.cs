using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Tool.Commands.Reader;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    private async Task<AgentConvertResult> ConvertIWorkAsync(string path, string outputPath, string format,
        bool overwrite, IWorkConversionOptions? options, long maximumInputBytes, long maximumOutputBytes,
        CancellationToken token) {
        string input = _pathPolicy.ResolveInput(path);
        if (Directory.Exists(input))
            throw new AgentUsageException("MCP iWork conversion requires a ZIP document file. Directory packages use the shared workflow provider API.");
        string destination = _pathPolicy.ResolveOutput(outputPath);
        ReaderToolPathSafety.EnsureDistinctFile(input, destination);
        string target = "." + format.ToLowerInvariant();
        if (!Path.GetExtension(destination).Equals(target, StringComparison.OrdinalIgnoreCase))
            throw new AgentUsageException("The output path extension must match the requested conversion format.");
        var runner = IWorkWorkflow.CreateRunner();
        string extension = Path.GetExtension(input).ToLowerInvariant();
        string route = runner.ConversionRoutes.SingleOrDefault(item => item.SourceExtensions.Contains(extension)
            && item.TargetExtension == target && item.Engine.Contains("IWork", StringComparison.Ordinal))?.Id
            ?? throw new AgentUsageException("Choose a Pages-to-DOCX, Numbers-to-XLSX, Keynote-to-PPTX, or iWork-to-PDF conversion.");
        AgentSourceRegistration source = _registry.Register(input, token);
        var request = new OfficeWorkflowRequest {
            InputPath = input, OutputPath = destination, Operation = OfficeWorkflowOperation.Convert,
            ConversionRouteId = route,
            ConflictPolicy = overwrite ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail,
            Limits = new OfficeWorkflowLimits { MaximumInputBytes = maximumInputBytes, MaximumOutputBytes = maximumOutputBytes },
            RegisteredConversionSettings = new IWorkWorkflowSettings { ConversionOptions = options ?? new() },
            PublicationGuard = new AgentConversionPublicationGuard(_pathPolicy, input)
        };
        OfficeWorkflowResult workflow = await runner.RunAsync(request, cancellationToken: token).ConfigureAwait(false);
        string hash = string.Empty;
        if (workflow.Succeeded) {
            await using var output = new FileStream(destination, FileMode.Open, FileAccess.Read, FileShare.Read,
                64 * 1024, useAsync: true);
            hash = Convert.ToHexString(await SHA256.HashDataAsync(output, token).ConfigureAwait(false)).ToLowerInvariant();
        }
        return new AgentConvertResult {
            SourceId = source.SourceId, SourcePath = source.Path, OutputPath = workflow.OutputPath ?? destination,
            Format = format.ToLowerInvariant(), LengthBytes = workflow.OutputBytes, Sha256 = hash,
            DiagnosticCount = workflow.Diagnostics.Count, Workflow = workflow,
            SourceSha256 = workflow.Diagnostics.FirstOrDefault(diagnostic => diagnostic.Code == "SourceSnapshot")?.Details.GetValueOrDefault("sha256")
        };
    }

    private sealed class AgentConversionPublicationGuard(AgentPathPolicy paths, string source) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (isDirectory || !string.Equals(paths.ResolveInput(source), source, StringComparison.Ordinal)
                || !string.Equals(paths.ResolveOutput(path), path, StringComparison.Ordinal)) return ValueTask.FromResult(false);
            ReaderToolPathSafety.EnsureDistinctFile(source, path);
            return ValueTask.FromResult(true);
        }
    }
}
