using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

/// <summary>Hashes a bounded seekable source without changing its caller-visible position.</summary>
internal static class EmailStoreSourceFingerprint {
    internal static string Compute(Stream stream, long expectedLength, long maximumBytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        long currentLength = stream.Length;
        if (currentLength > maximumBytes) {
            throw new EmailStoreLimitExceededException(
                nameof(EmailStoreReaderOptions.MaxInputBytes), currentLength, maximumBytes);
        }
        if (currentLength != expectedLength) {
            throw new InvalidDataException("The email-store source length changed after the session was opened.");
        }

        long position = stream.Position;
        try {
            stream.Position = 0;
            using IncrementalHash fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var buffer = new byte[64 * 1024];
            long totalRead = 0;
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                totalRead += read;
                if (totalRead > expectedLength || totalRead > maximumBytes) {
                    throw new InvalidDataException("The email-store source changed while its durable fingerprint was computed.");
                }
                fingerprint.AppendData(buffer, 0, read);
            }
            if (totalRead != expectedLength) {
                throw new InvalidDataException("The email-store source ended before its declared length while its durable fingerprint was computed.");
            }
            long finalLength = stream.Length;
            if (finalLength > maximumBytes) {
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxInputBytes), finalLength, maximumBytes);
            }
            if (finalLength != expectedLength) {
                throw new InvalidDataException("The email-store source changed while its durable fingerprint was computed.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            return EmailHashing.ToHexLower(fingerprint.GetHashAndReset());
        } finally {
            stream.Position = position;
        }
    }
}
