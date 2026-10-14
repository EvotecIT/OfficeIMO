using OfficeIMO.Reader;
using OfficeIMO.Tool.Commands.Reader;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    private static bool IsIWorkPackage(string path) => AgentSourceRegistry.IsIWorkPackagePath(path);

    // Native content and its identity must come from one Reader capture. An
    // earlier hash followed by a fresh read can label changed content with a stale id.
    private async Task<(AgentSourceRegistration Source, OfficeDocumentReadResult Document)> ReadDocumentSourceAsync(
        string path, AgentSourceRegistration? source, CancellationToken token, string? expectedSourceId = null) {
        long maximumInputBytes = source?.IWorkMaximumPackageBytes
            ?? (expectedSourceId == null ? null : _registry.GetRegisteredIWorkMaximumPackageBytes(expectedSourceId))
            ?? ReaderToolArguments.DefaultMaxInputBytes;
        if (source != null)
            return (source, await ReadDocumentAsync(path, token, maximumInputBytes: maximumInputBytes).ConfigureAwait(false));

        string physicalIdentity = OfficeImoToolPathSafety.GetPhysicalIdentityKey(path);
        bool directoryPackage = Directory.Exists(path);
        DateTime writeTime = directoryPackage ? Directory.GetLastWriteTimeUtc(path) : File.GetLastWriteTimeUtc(path);
        OfficeDocumentReadResult document;
        try {
            document = await ReadDocumentAsync(path, token, computeHashes: true,
                maximumInputBytes: maximumInputBytes).ConfigureAwait(false);
        } catch (IOException error) when (expectedSourceId != null && !directoryPackage) {
            throw new AgentUsageException("The registered iWork source could not be recaptured within its admitted policy. Inspect or search it again. " + error.Message);
        }
        if (!(directoryPackage ? Directory.Exists(path) : File.Exists(path))
            || !string.Equals(_pathPolicy.ResolveInput(path), path, StringComparison.Ordinal)
            || OfficeImoToolPathSafety.GetPhysicalIdentityKey(path) != physicalIdentity)
            throw new AgentUsageException("The source root changed during capture. Inspect or search it again.");
        string contentHash = document.Source.SourceHash
            ?? throw new InvalidDataException("The iWork Reader capture did not provide its source hash.");
        source = expectedSourceId == null
            ? _registry.RegisterIWorkSnapshot(path, directoryPackage, contentHash, writeTime, token, maximumInputBytes, document.Source.LengthBytes)
            : _registry.ResolveIWorkSnapshot(expectedSourceId, path, directoryPackage, contentHash, writeTime, token, maximumInputBytes, document.Source.LengthBytes);
        return (source, document);
    }
}
