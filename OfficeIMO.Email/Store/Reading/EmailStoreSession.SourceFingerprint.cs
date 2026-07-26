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
}
