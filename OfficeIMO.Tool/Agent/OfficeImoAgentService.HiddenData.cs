using OfficeIMO.Tool.Commands.HiddenData;
using OfficeIMO.Tool.Commands.Reader;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal Task<HiddenDataResultDto> InspectHiddenDataAsync(string path, CancellationToken cancellationToken = default) {
        string input = _pathPolicy.ResolveInput(path);
        EnsureHiddenDataSource(input);
        OfficeDocumentInfo info = OfficeDocument.Inspect(input, new OfficeInspectOptions { IncludeContent = false, IncludeHiddenData = true }, cancellationToken);
        return Task.FromResult(OfficeHiddenDataReportSerializer.Create(info.HiddenData!, input));
    }

    internal Task<HiddenDataResultDto> RemoveHiddenDataAsync(string path, string outputPath, IReadOnlyList<string>? kinds = null,
        string trackedChanges = "keep", IReadOnlyList<string>? keepProvenance = null, string? anonymousAuthor = null, bool removeInvalidatedSignatures = false, bool overwrite = false,
        CancellationToken cancellationToken = default) {
        string input = _pathPolicy.ResolveInput(path);
        EnsureHiddenDataSource(input);
        string destination = _pathPolicy.ResolveOutput(outputPath);
        ReaderToolPathSafety.EnsureDistinctFile(input, destination);
        if (!overwrite && File.Exists(destination)) {
            throw new AgentUsageException("Output file already exists. Choose a new path or explicitly enable overwrite.");
        }
        var options = new OfficeHiddenDataOptions();
        try {
            if (kinds is { Count: > 0 }) options.Kinds = HiddenDataArgumentParser.ParseKinds(kinds);
            options.TrackedChanges = HiddenDataArgumentParser.ParseTrackedChanges(trackedChanges);
            if (keepProvenance != null) HiddenDataArgumentParser.ApplyKeptProvenance(options, keepProvenance);
            if (anonymousAuthor != null) options.AnonymousAuthor = anonymousAuthor;
        } catch (ArgumentException exception) {
            throw new AgentUsageException(exception.Message, exception);
        }
        cancellationToken.ThrowIfCancellationRequested();
        try {
            OfficeHiddenDataReport report = OfficeDocument.RemoveHiddenData(input, destination, options,
                removeInvalidatedSignatures ? OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures : OfficeSignatureMutationPolicy.BlockSave);
            return Task.FromResult(OfficeHiddenDataReportSerializer.Create(report, input, destination));
        } catch (InvalidOperationException exception) {
            throw new AgentUsageException(exception.Message, exception);
        }
    }

    private static void EnsureHiddenDataSource(string input) {
        if (!File.Exists(input)) throw new FileNotFoundException("Input file was not found.", input);
        if (!OfficeDocument.SupportsHiddenData(input)) {
            throw new AgentUsageException("Hidden-data inspection supports Office, Visio, OpenDocument, EPUB, PDF and JPEG, PNG, WebP, GIF and TIFF image files.");
        }
    }
}
