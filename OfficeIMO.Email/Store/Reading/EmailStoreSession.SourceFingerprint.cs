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

        long expectedLength = SourceLength;
        long currentLength = _stream.Length;
        if (currentLength > _options.MaxInputBytes) {
            throw new EmailStoreLimitExceededException(
                nameof(EmailStoreReaderOptions.MaxInputBytes), currentLength, _options.MaxInputBytes);
        }
        if (currentLength != expectedLength) {
            throw new InvalidDataException("The email-store source length changed after the session was opened.");
        }

        long position = _stream.Position;
        try {
            _stream.Position = 0;
            using IncrementalHash fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var buffer = new byte[64 * 1024];
            long totalRead = 0;
            int read;
            while ((read = _stream.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                totalRead += read;
                if (totalRead > expectedLength || totalRead > _options.MaxInputBytes) {
                    throw new InvalidDataException("The email-store source changed while its durable fingerprint was computed.");
                }
                fingerprint.AppendData(buffer, 0, read);
            }
            if (totalRead != expectedLength) {
                throw new InvalidDataException("The email-store source ended before its declared length while its durable fingerprint was computed.");
            }
            long finalLength = _stream.Length;
            if (finalLength > _options.MaxInputBytes) {
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxInputBytes), finalLength, _options.MaxInputBytes);
            }
            if (finalLength != expectedLength) {
                throw new InvalidDataException("The email-store source changed while its durable fingerprint was computed.");
            }
            return EmailHashing.ToHexLower(fingerprint.GetHashAndReset());
        } finally {
            _stream.Position = position;
        }
    }
}
