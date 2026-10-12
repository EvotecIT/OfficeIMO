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
        bool directoryPackage = Directory.Exists(input);
        DateTime directoryWriteTime = directoryPackage ? Directory.GetLastWriteTimeUtc(input) : default;
        string destination = _pathPolicy.ResolveOutput(outputPath);
        ReaderToolPathSafety.EnsureDistinctFile(input, destination);
        if (directoryPackage) ReaderToolPathSafety.EnsureOutsideInput(input, destination);
        string target = "." + format.ToLowerInvariant();
        if (!Path.GetExtension(destination).Equals(target, StringComparison.OrdinalIgnoreCase))
            throw new AgentUsageException("The output path extension must match the requested conversion format.");
        var runner = IWorkWorkflow.CreateRunner();
        string extension = Path.GetExtension(input).ToLowerInvariant();
        string route = runner.ConversionRoutes.SingleOrDefault(item => item.SourceExtensions.Contains(extension)
            && item.TargetExtension == target && item.Engine.Contains("IWork", StringComparison.Ordinal))?.Id
            ?? throw new AgentUsageException("Choose a Pages-to-DOCX, Numbers-to-XLSX, Keynote-to-PPTX, or iWork-to-PDF conversion.");
        AgentSourceRegistration? source = directoryPackage ? null : _registry.Register(input, token);
        var settings = new IWorkWorkflowSettings { ConversionOptions = options ?? new() };
        var request = new OfficeWorkflowRequest {
            InputPath = input, OutputPath = destination, Operation = OfficeWorkflowOperation.Convert,
            ConversionRouteId = route,
            ConflictPolicy = overwrite ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail,
            Limits = new OfficeWorkflowLimits { MaximumInputBytes = maximumInputBytes, MaximumOutputBytes = maximumOutputBytes },
            RegisteredConversionSettings = settings,
            PublicationGuard = new AgentConversionPublicationGuard(_pathPolicy, input)
        };
        OfficeWorkflowResult workflow = await runner.RunAsync(request, cancellationToken: token).ConfigureAwait(false);
        if (directoryPackage && workflow.ConversionEvidence?.Facts.GetValueOrDefault("packageContentSha256") is string contentHash)
            // Result construction binds already captured evidence, including cancellation.
            // It performs no filesystem read and must not discard the workflow's result.
            source = _registry.RegisterIWorkDirectorySnapshot(input, contentHash, directoryWriteTime, CancellationToken.None,
                Math.Min(maximumInputBytes, settings.ReadOptions.MaximumPackageBytes));
        return new AgentConvertResult {
            SourceId = source?.SourceId ?? string.Empty, SourcePath = input, OutputPath = workflow.OutputPath ?? string.Empty,
            Format = format.ToLowerInvariant(), LengthBytes = workflow.OutputBytes, Sha256 = workflow.OutputSha256 ?? string.Empty,
            DiagnosticCount = workflow.Diagnostics.Count, Workflow = workflow,
            SourceSha256 = workflow.Diagnostics.FirstOrDefault(diagnostic => diagnostic.Code == "SourceSnapshot")?.Details.GetValueOrDefault("sha256")
        };
    }

    private sealed class AgentConversionPublicationGuard(AgentPathPolicy paths, string source) : IOfficeWorkflowPublicationGuard {
        private readonly string _sourceIdentity = OfficeImoToolPathSafety.GetPhysicalIdentityKey(source);
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (isDirectory || !string.Equals(paths.ResolveInput(source), source, StringComparison.Ordinal)
                || OfficeImoToolPathSafety.GetPhysicalIdentityKey(source) != _sourceIdentity
                || !string.Equals(paths.ResolveOutput(path), path, StringComparison.Ordinal)) return ValueTask.FromResult(false);
            ReaderToolPathSafety.EnsureDistinctFile(source, path);
            if (Directory.Exists(source)) ReaderToolPathSafety.EnsureOutsideInput(source, path);
            return ValueTask.FromResult(true);
        }
    }
}
