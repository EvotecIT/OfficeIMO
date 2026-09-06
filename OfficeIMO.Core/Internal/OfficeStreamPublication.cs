using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
    /// <summary>Publishes pre-serialized artifacts through storage-provider streams with read-back verification.</summary>
    /// <remarks>Stream providers do not offer atomic compare-and-replace. Callers must retain recoverable edits
    /// on any failure and disclose that concurrent changes or interrupted writes cannot be rolled back.</remarks>
    internal static class OfficeStreamPublication {
        private const string UncertainPublication = "OfficeIMO.Storage.DestinationMayHaveChanged";

        /// <summary>Indicates that a failed provider write may already have changed the destination.</summary>
        internal static bool MayHaveChangedDestination(Exception exception) => exception.Data.Contains(UncertainPublication);

        internal static async Task<string> WriteVerifiedAsync(
            Func<CancellationToken, Task<Stream>> openRead,
            Func<CancellationToken, Task<Stream>> openWrite,
            byte[] bytes,
            string? expectedFingerprint,
            long maximumBytes,
            CancellationToken cancellationToken) {
            if (openRead == null) throw new ArgumentNullException(nameof(openRead));
            if (openWrite == null) throw new ArgumentNullException(nameof(openWrite));
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (maximumBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            if (bytes.LongLength > maximumBytes) throw new InvalidDataException("The output exceeds the configured size limit.");
            cancellationToken.ThrowIfCancellationRequested();
            if (expectedFingerprint != null) {
                string current = await ReadFingerprintAsync(openRead, maximumBytes, cancellationToken).ConfigureAwait(false);
                if (!string.Equals(current, expectedFingerprint, StringComparison.OrdinalIgnoreCase)) {
                    throw new IOException("The source changed after it was opened. Save the edits to a different destination.");
                }
            }
            string outputFingerprint;
            using (SHA256 hash = SHA256.Create()) outputFingerprint = ToHex(hash.ComputeHash(bytes));
            cancellationToken.ThrowIfCancellationRequested();
            try {
                using (Stream destination = await openWrite(cancellationToken).ConfigureAwait(false)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    await OfficeStreamWriter.WriteAllBytesAsync(destination, bytes, cancellationToken).ConfigureAwait(false);
                }
                // A provider may commit its buffered output when the write stream is closed.
                string published = await ReadFingerprintAsync(openRead, maximumBytes, cancellationToken).ConfigureAwait(false);
                if (!string.Equals(published, outputFingerprint, StringComparison.Ordinal)) {
                    throw new IOException("The provider did not retain the complete output. Keep the edits and save to another destination.");
                }
            } catch (Exception error) {
                // Opening a provider's write stream may truncate before it returns or throws.
                error.Data[UncertainPublication] = true;
                throw;
            }
            return outputFingerprint;
        }

        internal static async Task<string> ReadFingerprintAsync(
            Func<CancellationToken, Task<Stream>> openRead, long maximumBytes, CancellationToken cancellationToken) {
            if (openRead == null) throw new ArgumentNullException(nameof(openRead));
            if (maximumBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            cancellationToken.ThrowIfCancellationRequested();
            using (Stream source = await openRead(cancellationToken).ConfigureAwait(false))
            using (SHA256 hash = SHA256.Create()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!source.CanRead) throw new ArgumentException("The provider stream must be readable.");
                if (source.CanSeek) {
                    if (source.Length > maximumBytes) throw new InvalidDataException("The input exceeds the configured size limit.");
                    source.Position = 0;
                }
                byte[] buffer = new byte[81920];
                long total = 0;
                int read;
                while ((read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                    total = checked(total + read);
                    if (total > maximumBytes) throw new InvalidDataException("The input exceeds the configured size limit.");
                    hash.TransformBlock(buffer, 0, read, buffer, 0);
                }
                cancellationToken.ThrowIfCancellationRequested();
                hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                return ToHex(hash.Hash!);
            }
        }

        internal static async Task VerifyFingerprintAsync(Func<CancellationToken, Task<Stream>> openRead,
            string expectedFingerprint, long maximumBytes, CancellationToken cancellationToken) {
            string current = await ReadFingerprintAsync(openRead, maximumBytes, cancellationToken).ConfigureAwait(false);
            if (!string.Equals(current, expectedFingerprint, StringComparison.OrdinalIgnoreCase)) {
                throw new IOException("The input changed while the workflow was running. No output was published.");
            }
        }

        private static string ToHex(byte[] bytes) => BitConverter.ToString(bytes).Replace("-", string.Empty);
    }
}
