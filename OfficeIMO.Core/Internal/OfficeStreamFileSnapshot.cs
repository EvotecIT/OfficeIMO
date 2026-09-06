using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
    /// <summary>Stages a bounded provider stream for engines that require a seekable local input.</summary>
    internal sealed class OfficeStreamFileSnapshot : IDisposable {
        internal const string CleanupFailureDataKey = "OfficeIMO.Storage.StagingCleanupFailed";
        private readonly string _directory;
        private FileStream? _lease;
        private bool _disposed;

        private OfficeStreamFileSnapshot(string directory, string path, long length, string fingerprint, FileStream lease) {
            _directory = directory;
            FilePath = path;
            Length = length;
            Fingerprint = fingerprint;
            _lease = lease;
        }

        internal string FilePath { get; }
        internal long Length { get; }
        internal string Fingerprint { get; }

        internal static async Task<OfficeStreamFileSnapshot> CaptureAsync(
            Func<CancellationToken, Task<Stream>> openRead, string extension, long maximumBytes,
            string? expectedFingerprint, CancellationToken token) {
            if (openRead == null) throw new ArgumentNullException(nameof(openRead));
            if (maximumBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            if (extension == null || extension.Length < 2 || extension.Length > 17 || extension[0] != '.') {
                throw new ArgumentException("A short file extension is required.", nameof(extension));
            }
            for (int index = 1; index < extension.Length; index++) {
                if (!char.IsLetterOrDigit(extension[index])) throw new ArgumentException("The file extension must contain only letters and digits.", nameof(extension));
            }
            token.ThrowIfCancellationRequested();
            string directory = OfficeTemporaryDirectory.Create("officeimo-input-");
            string path = Path.Combine(directory, "input" + extension);
            FileStream? lease = null;
            try {
                long total = 0;
                string fingerprint;
                using (Stream source = await openRead(token).ConfigureAwait(false))
                using (FileStream output = OfficeTemporaryFile.CreateAtPath(path, 81920, FileOptions.Asynchronous | FileOptions.SequentialScan))
                using (SHA256 hash = SHA256.Create()) {
                    token.ThrowIfCancellationRequested();
                    if (!source.CanRead) throw new ArgumentException("The provider stream must be readable.");
                    if (source.CanSeek) {
                        if (source.Length > maximumBytes) throw new InvalidDataException("The input exceeds the configured size limit.");
                        source.Position = 0;
                    }
                    byte[] buffer = new byte[81920];
                    int read;
                    while ((read = await source.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false)) > 0) {
                        token.ThrowIfCancellationRequested();
                        total = checked(total + read);
                        if (total > maximumBytes) throw new InvalidDataException("The input exceeds the configured size limit.");
                        hash.TransformBlock(buffer, 0, read, buffer, 0);
                        await output.WriteAsync(buffer, 0, read, token).ConfigureAwait(false);
                    }
                    token.ThrowIfCancellationRequested();
                    hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                    fingerprint = BitConverter.ToString(hash.Hash!).Replace("-", string.Empty);
                    if (expectedFingerprint != null && !string.Equals(expectedFingerprint, fingerprint, StringComparison.OrdinalIgnoreCase)) {
                        throw new IOException("The input changed after it was selected. Select it again before running the workflow.");
                    }
                    await output.FlushAsync(token).ConfigureAwait(false);
                }
                lease = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                token.ThrowIfCancellationRequested();
                var result = new OfficeStreamFileSnapshot(directory, path, total, fingerprint, lease);
                lease = null;
                return result;
            } catch (Exception error) {
                lease?.Dispose();
                try { File.Delete(path); Directory.Delete(directory, recursive: false); }
                catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException) {
                    error.Data[CleanupFailureDataKey] = directory;
                }
                throw;
            }
        }

        internal async Task VerifySourceAsync(Func<CancellationToken, Task<Stream>> openRead,
            long maximumBytes, CancellationToken token) {
            if (_disposed) throw new ObjectDisposedException(nameof(OfficeStreamFileSnapshot));
            await OfficeStreamPublication.VerifyFingerprintAsync(openRead, Fingerprint, maximumBytes, token).ConfigureAwait(false);
        }

        public void Dispose() {
            if (_disposed) return;
            _lease?.Dispose();
            _lease = null;
            File.Delete(FilePath);
            Directory.Delete(_directory, recursive: false);
            _disposed = true;
        }
    }
}
