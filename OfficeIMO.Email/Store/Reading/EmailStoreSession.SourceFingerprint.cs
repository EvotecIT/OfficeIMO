using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Computes a stable metadata fingerprint from the session's already bounded, link-safe source catalog.
    /// Message bodies and attachment payloads are not materialized.
    /// </summary>
    public string GetCatalogFingerprint(CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        if (_backend is MailboxDirectoryStoreSessionBackend directory) {
            return directory.GetCatalogFingerprint(cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();
        string value = string.Join(
            "|",
            Format.ToString(),
            SourceLength.ToString(System.Globalization.CultureInfo.InvariantCulture),
            DisplayName ?? string.Empty,
            Folders.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return EmailHashing.ComputeSha256HexLower(value);
    }

    /// <summary>
    /// Computes a SHA-256 fingerprint over the complete persisted source. This intentionally performs source I/O
    /// so a continuation created in another process cannot be accepted after same-length content was changed.
    /// </summary>
    public string GetDurableSourceFingerprint(CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        if (_backend is MailboxDirectoryStoreSessionBackend directory) {
            return directory.GetContentFingerprint(cancellationToken);
        }

        long position = _stream.Position;
        try {
            _stream.Position = 0;
            using IncrementalHash fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var buffer = new byte[64 * 1024];
            int read;
            while ((read = _stream.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                fingerprint.AppendData(buffer, 0, read);
            }
            return EmailHashing.ToHexLower(fingerprint.GetHashAndReset());
        } finally {
            _stream.Position = position;
        }
    }
}
