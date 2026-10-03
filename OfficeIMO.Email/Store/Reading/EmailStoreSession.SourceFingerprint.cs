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
        cancellationToken.ThrowIfCancellationRequested();
        if (_snapshotFingerprint != null) return _snapshotFingerprint;
        if (_backend is MailboxDirectoryStoreSessionBackend directory) {
            return directory.GetContentFingerprint(cancellationToken);
        }

        return EmailStoreSourceFingerprint.Compute(_stream, SourceLength, _options.MaxInputBytes, cancellationToken);
    }
}
