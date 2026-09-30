using OfficeIMO.Email.Store;
using OfficeIMO.Email.Data;
using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Tool.Agent;

internal sealed class AgentSourceRegistry {
    private const string EmailDataPrefix = "officeimo-email-data:";
    private const string EmailStorePrefix = "officeimo-email-store:";
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

    internal AgentSourceRegistration RegisterEmailData(string path, EmailDataOpenOptions openOptions,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(openOptions);
        AgentSourceRegistration registration = Create(path, cancellationToken, openOptions);
        _sources[registration.SourceId] = registration;
        return registration;
    }

    internal AgentSourceRegistration RegisterEmailStore(string path, EmailStoreReaderOptions storeOptions,
        CancellationToken cancellationToken = default) {
        var options = new EmailDataOpenOptions(store: storeOptions, expectedKind: EmailDataArtifactKind.Store);
        AgentSourceRegistration registration = Create(path, cancellationToken, options, emailContentSearch: true);
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
        AgentSourceRegistration current = Create(registered.Path, cancellationToken, registered.EmailDataOptions, registered.EmailContentSearch);
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
        bool contentSearch = sourceId.StartsWith(EmailStorePrefix, StringComparison.Ordinal);
        EmailDataOpenOptions? dataOptions = _sources.TryGetValue(sourceId, out var registered)
            ? registered.EmailDataOptions : contentSearch
                ? new EmailDataOpenOptions(store: OfficeImoAgentService.CreateEmailStoreOptions(), expectedKind: EmailDataArtifactKind.Store)
                : sourceId.StartsWith(EmailDataPrefix, StringComparison.Ordinal) ? new EmailDataInspectionOptions().OpenOptions : null;
        AgentSourceRegistration current = Create(path, cancellationToken, dataOptions, contentSearch);
        if (!string.Equals(current.SourceId, sourceId, StringComparison.Ordinal)) {
            throw new AgentUsageException(
                "The supplied path does not match the source id, or the source changed. Search it again.");
        }
        _sources[sourceId] = current;
        return current;
    }

    private static AgentSourceRegistration Create(
        string path,
        CancellationToken cancellationToken, EmailDataOpenOptions? dataOptions = null, bool emailContentSearch = false) {
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
            cancellationToken, dataOptions);
        return new AgentSourceRegistration(
            (emailContentSearch ? EmailStorePrefix : dataOptions == null ? "officeimo:" : EmailDataPrefix) + hash.Substring(0, 24),
            fullPath,
            isDirectory,
            length,
            lastWriteUtc, dataOptions, emailContentSearch);
    }

    private static string CreateHash(
        string fullPath,
        bool isDirectory,
        long? length,
        DateTime lastWriteUtc,
        CancellationToken cancellationToken, EmailDataOpenOptions? dataOptions) {
        using IncrementalHash fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        Append(fingerprint, fullPath);
        Append(fingerprint, isDirectory ? "directory" : "file");
        Append(fingerprint, length?.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Append(fingerprint, lastWriteUtc.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (dataOptions != null) {
            // Identity discovery uses the same bounded artifact owners as the inspector.
            using var opened = EmailDataArtifact.Open(fullPath, dataOptions, cancellationToken);
            Append(fingerprint, opened.Kind.ToString());
            if (opened.Store != null) Append(fingerprint, opened.Store.GetDurableSourceFingerprint(cancellationToken));
            else if (opened.AddressBook != null) Append(fingerprint, opened.AddressBook.GetDurableSourceFingerprint(cancellationToken));
            else {
                long maximum = opened.Email != null ? dataOptions.Email.MaxInputBytes : dataOptions.ContentLines.MaxInputBytes;
                using var input = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var buffer = new byte[64 * 1024]; long total = 0; int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) != 0) {
                    cancellationToken.ThrowIfCancellationRequested(); total = checked(total + read);
                    if (total > maximum) throw new IOException("The mail-data source exceeds its bounded identity-read policy.");
                    fingerprint.AppendData(buffer, 0, read);
                }
            }
        } else if (isDirectory) {
            using EmailStoreSession session = EmailStoreSession.Open(
                fullPath,
                cancellationToken: cancellationToken);
            Append(fingerprint, session.GetCatalogFingerprint(cancellationToken));
        }
        cancellationToken.ThrowIfCancellationRequested();
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
    DateTime LastWriteUtc,
    EmailDataOpenOptions? EmailDataOptions = null,
    bool EmailContentSearch = false);
