using OfficeIMO.Email.Store;
using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Tool.Agent;

internal sealed class AgentSourceRegistry {
    private static readonly byte[] FingerprintSeparator = new byte[] { 0 };
    private readonly ConcurrentDictionary<string, AgentSourceRegistration> _sources =
        new(StringComparer.Ordinal);

    internal AgentSourceRegistration Register(
        string path,
        CancellationToken cancellationToken = default) {
        AgentSourceRegistration registration = Create(path, cancellationToken);
        _sources[registration.SourceId] = registration;
        return registration;
    }

    internal AgentSourceRegistration Resolve(
        string sourceId,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(sourceId) ||
            !_sources.TryGetValue(sourceId, out AgentSourceRegistration? registered)) {
            throw new AgentUsageException(
                "Unknown source id. Inspect or search the source again before fetching content.");
        }
        AgentSourceRegistration current = Create(registered.Path, cancellationToken);
        if (!string.Equals(current.SourceId, sourceId, StringComparison.Ordinal)) {
            _sources.TryRemove(sourceId, out _);
            throw new AgentUsageException(
                "The source changed after the id was issued. Inspect or search it again.");
        }
        return current;
    }

    internal AgentSourceRegistration Resolve(
        string sourceId,
        string path,
        CancellationToken cancellationToken = default) {
        AgentSourceRegistration current = Create(path, cancellationToken);
        if (!string.Equals(current.SourceId, sourceId, StringComparison.Ordinal)) {
            throw new AgentUsageException(
                "The supplied path does not match the source id, or the source changed. Search it again.");
        }
        _sources[sourceId] = current;
        return current;
    }

    private static AgentSourceRegistration Create(
        string path,
        CancellationToken cancellationToken) {
        string fullPath = OfficeImoToolPathSafety.ResolveExistingLinks(path);
        bool isDirectory = Directory.Exists(fullPath);
        long? length = isDirectory ? null : new FileInfo(fullPath).Length;
        DateTime lastWriteUtc = isDirectory
            ? Directory.GetLastWriteTimeUtc(fullPath)
            : File.GetLastWriteTimeUtc(fullPath);
        string hash = CreateHash(
            fullPath,
            isDirectory,
            length,
            lastWriteUtc,
            cancellationToken);
        return new AgentSourceRegistration(
            "officeimo:" + hash.Substring(0, 24),
            fullPath,
            isDirectory,
            length,
            lastWriteUtc);
    }

    private static string CreateHash(
        string fullPath,
        bool isDirectory,
        long? length,
        DateTime lastWriteUtc,
        CancellationToken cancellationToken) {
        using IncrementalHash fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        Append(fingerprint, fullPath);
        Append(fingerprint, isDirectory ? "directory" : "file");
        Append(fingerprint, length?.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Append(fingerprint, lastWriteUtc.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (isDirectory) {
            using EmailStoreSession session = EmailStoreSession.Open(
                fullPath,
                cancellationToken: cancellationToken);
            Append(fingerprint, session.GetCatalogFingerprint(cancellationToken));
        }
        return Convert.ToHexString(fingerprint.GetHashAndReset()).ToLowerInvariant();
    }

    private static void Append(IncrementalHash fingerprint, string? value) {
        fingerprint.AppendData(Encoding.UTF8.GetBytes(value ?? string.Empty));
        fingerprint.AppendData(FingerprintSeparator);
    }
}

internal sealed record AgentSourceRegistration(
    string SourceId,
    string Path,
    bool IsDirectory,
    long? LengthBytes,
    DateTime LastWriteUtc);