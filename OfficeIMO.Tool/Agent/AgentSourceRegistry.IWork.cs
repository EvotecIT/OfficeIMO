using OfficeIMO.IWork;
using OfficeIMO.Tool.Commands.Reader;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class AgentSourceRegistry {
    internal AgentSourceRegistration RegisterIWorkFile(string path, long maximumPackageBytes, CancellationToken token) {
        if (!IsIWorkPackagePath(path)) throw new ArgumentException("An iWork package path is required.", nameof(path));
        if (maximumPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPackageBytes));
        var source = Create(path, token, iworkMaximumPackageBytes: maximumPackageBytes);
        _sources[source.SourceId] = source;
        return source;
    }

    internal long? GetRegisteredIWorkMaximumPackageBytes(string sourceId) =>
        !string.IsNullOrWhiteSpace(sourceId) && _sources.TryGetValue(sourceId, out var source)
            ? source.IWorkMaximumPackageBytes : null;

    // Reader and conversion supply the format owner's captured-content hash.
    // Identity construction never reopens the package after that capture.
    internal AgentSourceRegistration RegisterIWorkDirectorySnapshot(string path, string contentHash,
        DateTime capturedLastWriteUtc, CancellationToken token, long maximumPackageBytes = ReaderToolArguments.DefaultMaxInputBytes) {
        var source = CreateIWorkDirectorySnapshot(path, contentHash, capturedLastWriteUtc, token, maximumPackageBytes);
        _sources[source.SourceId] = source;
        return source;
    }

    internal AgentSourceRegistration ResolveIWorkDirectorySnapshot(string sourceId, string path, string contentHash,
        DateTime capturedLastWriteUtc, CancellationToken token, long maximumPackageBytes = ReaderToolArguments.DefaultMaxInputBytes) {
        var source = CreateIWorkDirectorySnapshot(path, contentHash, capturedLastWriteUtc, token, maximumPackageBytes);
        if (!string.Equals(source.SourceId, sourceId, StringComparison.Ordinal))
            throw new AgentUsageException("The supplied path does not match the source id, or the source changed. Search it again.");
        _sources[sourceId] = source;
        return source;
    }

    private static AgentSourceRegistration CreateIWorkDirectorySnapshot(string path, string contentHash,
        DateTime capturedLastWriteUtc, CancellationToken token, long maximumPackageBytes) {
        if (!IsIWorkPackagePath(path) || contentHash.Length != 64 || !contentHash.All(Uri.IsHexDigit))
            throw new ArgumentException("A captured iWork package-content SHA-256 is required.");
        if (maximumPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPackageBytes));
        // Path was authorized before capture. Do not resolve it again here: a failed
        // workflow may have observed a replaced, moved or revoked source root.
        string hash = CreateHash(path, true, null, capturedLastWriteUtc, token, null, contentHash.ToLowerInvariant());
        return new AgentSourceRegistration("officeimo:" + hash.Substring(0, 24), path, true, null, capturedLastWriteUtc,
            IWorkMaximumPackageBytes: maximumPackageBytes);
    }

    internal static bool IsIWorkPackagePath(string path) => Path.GetExtension(path).ToLowerInvariant() is ".pages" or ".numbers" or ".key";

    private static string ComputeIWorkContentHash(string path, long? maximumPackageBytes, CancellationToken token) =>
        IWorkSourceDocument.Open(path, new IWorkReadOptions {
            MaximumPackageBytes = maximumPackageBytes ?? ReaderToolArguments.DefaultMaxInputBytes,
            PreserveSourceRecords = false
        }, token).ComputePackageContentHash();
}
