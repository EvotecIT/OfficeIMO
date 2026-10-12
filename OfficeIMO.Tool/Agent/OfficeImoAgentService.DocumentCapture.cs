using OfficeIMO.Reader;
using OfficeIMO.Tool.Commands.Reader;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    private static bool IsIWorkDirectory(string path) =>
        Directory.Exists(path) && AgentSourceRegistry.IsIWorkPackagePath(path);

    // Directory content and its identity must come from one Reader capture. An
    // earlier hash followed by a fresh read can label changed content with a stale id.
    private async Task<(AgentSourceRegistration Source, OfficeDocumentReadResult Document)> ReadDocumentSourceAsync(
        string path, AgentSourceRegistration? source, CancellationToken token, string? expectedSourceId = null) {
        long maximumInputBytes = source?.IWorkMaximumPackageBytes
            ?? (expectedSourceId == null ? null : _registry.GetRegisteredIWorkMaximumPackageBytes(expectedSourceId))
            ?? ReaderToolArguments.DefaultMaxInputBytes;
        if (source != null)
            return (source, await ReadDocumentAsync(path, token, maximumInputBytes: maximumInputBytes).ConfigureAwait(false));

        string physicalIdentity = OfficeImoToolPathSafety.GetPhysicalIdentityKey(path);
        DateTime writeTime = Directory.GetLastWriteTimeUtc(path);
        OfficeDocumentReadResult document = await ReadDocumentAsync(path, token, computeHashes: true,
            maximumInputBytes: maximumInputBytes).ConfigureAwait(false);
        if (!Directory.Exists(path) || !string.Equals(_pathPolicy.ResolveInput(path), path, StringComparison.Ordinal)
            || OfficeImoToolPathSafety.GetPhysicalIdentityKey(path) != physicalIdentity)
            throw new AgentUsageException("The source root changed during capture. Inspect or search it again.");
        string contentHash = document.Source.SourceHash
            ?? throw new InvalidDataException("The iWork Reader capture did not provide its package-content hash.");
        source = expectedSourceId == null
            ? _registry.RegisterIWorkDirectorySnapshot(path, contentHash, writeTime, token, maximumInputBytes)
            : _registry.ResolveIWorkDirectorySnapshot(expectedSourceId, path, contentHash, writeTime, token, maximumInputBytes);
        return (source, document);
    }
}
