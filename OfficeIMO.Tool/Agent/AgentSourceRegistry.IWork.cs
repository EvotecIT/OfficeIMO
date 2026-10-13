using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Tool.Commands.Reader;
using System.Security.Cryptography;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class AgentSourceRegistry {
    internal long? GetRegisteredIWorkMaximumPackageBytes(string sourceId) =>
        !string.IsNullOrWhiteSpace(sourceId) && _sources.TryGetValue(sourceId, out var source)
            ? source.IWorkMaximumPackageBytes : null;

    // Reader and conversion supply the same capture hash: raw bytes for files,
    // package contents for directory bundles.
    // Identity construction never reopens the package after that capture.
    internal AgentSourceRegistration RegisterIWorkSnapshot(string path, bool directoryPackage, string contentHash,
        DateTime capturedLastWriteUtc, CancellationToken token, long maximumPackageBytes = ReaderToolArguments.DefaultMaxInputBytes,
        long? capturedLengthBytes = null) {
        var source = CreateIWorkSnapshot(path, directoryPackage, contentHash, capturedLastWriteUtc, token,
            maximumPackageBytes, capturedLengthBytes);
        _sources[source.SourceId] = source;
        return source;
    }

    internal AgentSourceRegistration ResolveIWorkSnapshot(string sourceId, string path, bool directoryPackage, string contentHash,
        DateTime capturedLastWriteUtc, CancellationToken token, long maximumPackageBytes = ReaderToolArguments.DefaultMaxInputBytes,
        long? capturedLengthBytes = null) {
        var source = CreateIWorkSnapshot(path, directoryPackage, contentHash, capturedLastWriteUtc, token,
            maximumPackageBytes, capturedLengthBytes);
        if (!string.Equals(source.SourceId, sourceId, StringComparison.Ordinal))
            throw new AgentUsageException("The supplied path does not match the source id, or the source changed. Search it again.");
        _sources[sourceId] = source;
        return source;
    }

    private static AgentSourceRegistration CreateIWorkSnapshot(string path, bool directoryPackage, string contentHash,
        DateTime capturedLastWriteUtc, CancellationToken token, long maximumPackageBytes, long? capturedLengthBytes) {
        if (!IsIWorkPackagePath(path) || contentHash.Length != 64 || !contentHash.All(Uri.IsHexDigit))
            throw new ArgumentException("A captured iWork source SHA-256 is required.");
        if (maximumPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPackageBytes));
        // Path was authorized before capture. Do not resolve it again here: a failed
        // workflow may have observed a replaced, moved or revoked source root.
        string hash = CreateHash(path, directoryPackage, capturedLengthBytes, capturedLastWriteUtc, token, null, contentHash.ToLowerInvariant());
        return new AgentSourceRegistration("officeimo:" + hash.Substring(0, 24), path, directoryPackage,
            directoryPackage ? null : capturedLengthBytes, capturedLastWriteUtc,
            IWorkMaximumPackageBytes: maximumPackageBytes);
    }

    internal static bool IsIWorkPackagePath(string path) => Path.GetExtension(path).ToLowerInvariant() is ".pages" or ".numbers" or ".key";

    private static string ComputeIWorkSourceHash(string path, bool directoryPackage, long? maximumPackageBytes, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        long maximum = maximumPackageBytes ?? ReaderToolArguments.DefaultMaxInputBytes;
        if (directoryPackage)
            return IWorkSourceDocument.Open(path, new IWorkReadOptions {
                MaximumPackageBytes = maximum, PreserveSourceRecords = false
            }, token).ComputePackageContentHash();
        using var input = IWorkContainerReader.OpenPackageFileForRead(path);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[64 * 1024]; long total = 0;
        while (true) {
            token.ThrowIfCancellationRequested();
            int read = input.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maximum) throw new AgentUsageException("The iWork source exceeds its admitted identity-read limit. Inspect or search it again.");
            hash.AppendData(buffer, 0, read);
        }
        token.ThrowIfCancellationRequested();
        return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
    }
}
