using OfficeIMO.IWork;
using OfficeIMO.Tool.Commands.Reader;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class AgentSourceRegistry {
    // Conversion supplies the format owner's captured-content hash, so source ids
    // refer to the converted snapshot without an additional post-publication read.
    internal AgentSourceRegistration RegisterIWorkDirectorySnapshot(string path, string contentHash,
        DateTime capturedLastWriteUtc, CancellationToken token) {
        if (!IsIWorkPackagePath(path) || contentHash.Length != 64 || !contentHash.All(Uri.IsHexDigit))
            throw new ArgumentException("A captured iWork package-content SHA-256 is required.");
        // Path was authorized before capture. Do not resolve it again here: a failed
        // workflow may have observed a replaced, moved or revoked source root.
        string hash = CreateHash(path, true, null, capturedLastWriteUtc, token, null, contentHash);
        var source = new AgentSourceRegistration("officeimo:" + hash.Substring(0, 24), path, true, null, capturedLastWriteUtc);
        _sources[source.SourceId] = source;
        return source;
    }

    internal static bool IsIWorkPackagePath(string path) => Path.GetExtension(path).ToLowerInvariant() is ".pages" or ".numbers" or ".key";

    private static string ComputeIWorkContentHash(string path, CancellationToken token) =>
        IWorkSourceDocument.Open(path, new IWorkReadOptions {
            MaximumPackageBytes = ReaderToolArguments.DefaultMaxInputBytes,
            PreserveSourceRecords = false
        }, token).ComputePackageContentHash();
}
