using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Commits completed Office files without exposing a partially written destination.
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficeFileCommit {
        /// <summary>Controls whether an existing destination may be replaced.</summary>
        public enum ConflictPolicy {
            /// <summary>Fails when the destination already exists.</summary>
            FailIfExists,
            /// <summary>Atomically replaces the destination when it exists.</summary>
            Replace
        }

        /// <summary>Produces a file in the destination directory and atomically commits it.</summary>
        public static void Write(string targetPath, Action<Stream> writer, ConflictPolicy conflictPolicy = ConflictPolicy.Replace) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(writer);
#else
            if (writer == null) throw new ArgumentNullException(nameof(writer));
#endif

            EnsureTargetDirectory(targetPath);
            string temporaryPath = CreateTemporaryPath(targetPath);
            try {
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) {
                    writer(stream);
                    stream.Flush();
                }

                CommitTemporaryFile(temporaryPath, targetPath, conflictPolicy);
                temporaryPath = string.Empty;
            } finally {
                DeleteIfExists(temporaryPath);
            }
        }

        /// <summary>Atomically writes a completed byte array to a destination path.</summary>
        public static void WriteAllBytes(string targetPath, byte[] bytes, ConflictPolicy conflictPolicy = ConflictPolicy.Replace) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(bytes);
#else
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
#endif
            Write(targetPath, stream => stream.Write(bytes, 0, bytes.Length), conflictPolicy);
        }

        /// <summary>Writes completed bytes to a same-directory staging file for a later atomic commit.</summary>
        public static string StageAllBytes(string targetPath, byte[] bytes) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(bytes);
#else
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
#endif
            EnsureTargetDirectory(targetPath);
            string fullTargetPath = GetFullTargetPath(targetPath);
            string temporaryPath = CreateTemporaryPath(fullTargetPath);
            try {
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) {
                    stream.Write(bytes, 0, bytes.Length);
                    stream.Flush();
                }

                return temporaryPath;
            } catch {
                DeleteIfExists(temporaryPath);
                throw;
            }
        }

        /// <summary>Asynchronously writes a completed byte array and atomically commits it.</summary>
        public static async Task WriteAllBytesAsync(
            string targetPath,
            byte[] bytes,
            ConflictPolicy conflictPolicy = ConflictPolicy.Replace,
            CancellationToken cancellationToken = default) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(bytes);
#else
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
#endif

            EnsureTargetDirectory(targetPath);
            string temporaryPath = CreateTemporaryPath(targetPath);
            try {
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 8192, FileOptions.Asynchronous)) {
#if NET6_0_OR_GREATER
                    await stream.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
                    await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
                    await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                }

                cancellationToken.ThrowIfCancellationRequested();
                CommitTemporaryFile(temporaryPath, targetPath, conflictPolicy);
                temporaryPath = string.Empty;
            } finally {
                DeleteIfExists(temporaryPath);
            }
        }

        /// <summary>Asynchronously writes completed bytes to a same-directory staging file.</summary>
        public static async Task<string> StageAllBytesAsync(
            string targetPath,
            byte[] bytes,
            CancellationToken cancellationToken = default) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(bytes);
#else
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
#endif
            cancellationToken.ThrowIfCancellationRequested();
            EnsureTargetDirectory(targetPath);
            string fullTargetPath = GetFullTargetPath(targetPath);
            string temporaryPath = CreateTemporaryPath(fullTargetPath);
            try {
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 8192, FileOptions.Asynchronous)) {
#if NET6_0_OR_GREATER
                    await stream.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
                    await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
                    await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                }

                cancellationToken.ThrowIfCancellationRequested();
                return temporaryPath;
            } catch {
                DeleteIfExists(temporaryPath);
                throw;
            }
        }

        /// <summary>Produces a file directly, asynchronously flushes it, and atomically commits it.</summary>
        public static async Task WriteAsync(
            string targetPath,
            Action<Stream> writer,
            ConflictPolicy conflictPolicy = ConflictPolicy.Replace,
            CancellationToken cancellationToken = default) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(writer);
#else
            if (writer == null) throw new ArgumentNullException(nameof(writer));
#endif
            cancellationToken.ThrowIfCancellationRequested();
            EnsureTargetDirectory(targetPath);
            string temporaryPath = CreateTemporaryPath(targetPath);
            try {
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 8192, FileOptions.Asynchronous)) {
                    writer(stream);
                    cancellationToken.ThrowIfCancellationRequested();
                    await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                }

                cancellationToken.ThrowIfCancellationRequested();
                CommitTemporaryFile(temporaryPath, targetPath, conflictPolicy);
                temporaryPath = string.Empty;
            } finally {
                DeleteIfExists(temporaryPath);
            }
        }

        /// <summary>Produces a file asynchronously and atomically commits it.</summary>
        public static async Task WriteAsync(
            string targetPath,
            Func<Stream, CancellationToken, Task> writer,
            ConflictPolicy conflictPolicy = ConflictPolicy.Replace,
            CancellationToken cancellationToken = default) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(writer);
#else
            if (writer == null) throw new ArgumentNullException(nameof(writer));
#endif
            cancellationToken.ThrowIfCancellationRequested();
            EnsureTargetDirectory(targetPath);
            string temporaryPath = CreateTemporaryPath(targetPath);
            try {
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite,
                           FileShare.None, 8192, FileOptions.Asynchronous)) {
                    await writer(stream, cancellationToken).ConfigureAwait(false);
                    cancellationToken.ThrowIfCancellationRequested();
                    await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                }

                cancellationToken.ThrowIfCancellationRequested();
                CommitTemporaryFile(temporaryPath, targetPath, conflictPolicy);
                temporaryPath = string.Empty;
            } finally {
                DeleteIfExists(temporaryPath);
            }
        }

        /// <summary>Creates a same-directory temporary path suitable for an atomic commit.</summary>
        public static string CreateTemporaryPath(string targetPath) {
            string fullTargetPath = GetFullTargetPath(targetPath);
            string? directory = Path.GetDirectoryName(fullTargetPath);
            if (string.IsNullOrEmpty(directory)) {
                directory = Directory.GetCurrentDirectory();
            }

            string fileName = Path.GetFileName(fullTargetPath);
            return Path.Combine(directory, $".{fileName}.{Guid.NewGuid():N}.tmp");
        }

        /// <summary>Creates a same-directory staging path that preserves the destination extension.</summary>
        public static string CreateStagingPath(string targetPath) {
            string fullTargetPath = GetFullTargetPath(targetPath);
            string? directory = Path.GetDirectoryName(fullTargetPath);
            if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();

            string extension = Path.GetExtension(fullTargetPath);
            string fileName = Path.GetFileNameWithoutExtension(fullTargetPath);
            return Path.Combine(directory, $".{fileName}.{Guid.NewGuid():N}{extension}");
        }

        /// <summary>Commits a completed temporary file to its destination.</summary>
        public static void CommitTemporaryFile(
            string temporaryPath,
            string targetPath,
            ConflictPolicy conflictPolicy = ConflictPolicy.Replace) {
            CommitTemporaryFileCore(temporaryPath, targetPath, conflictPolicy,
                allowNonAtomicReplacementFallback: true);
        }

        /// <summary>
        /// Commits a completed temporary file without falling back to a replacement sequence that
        /// temporarily removes an existing destination pathname.
        /// </summary>
        public static void CommitTemporaryFileAtomically(
            string temporaryPath,
            string targetPath,
            ConflictPolicy conflictPolicy = ConflictPolicy.Replace) {
            CommitTemporaryFileCore(temporaryPath, targetPath, conflictPolicy,
                allowNonAtomicReplacementFallback: false);
        }

        private static void CommitTemporaryFileCore(
            string temporaryPath,
            string targetPath,
            ConflictPolicy conflictPolicy,
            bool allowNonAtomicReplacementFallback) {
            if (string.IsNullOrWhiteSpace(temporaryPath)) throw new ArgumentException("Temporary path cannot be empty.", nameof(temporaryPath));

            string fullTargetPath = GetFullTargetPath(targetPath);
            if (conflictPolicy == ConflictPolicy.FailIfExists) {
                if (!TryMoveIfAbsent(temporaryPath, fullTargetPath)) {
                    throw new IOException($"Destination file '{fullTargetPath}' already exists.");
                }
                return;
            }

            EnsureDestinationWritable(fullTargetPath);

            if (!File.Exists(fullTargetPath)) {
                if (TryMoveIfAbsent(temporaryPath, fullTargetPath, waitForClaim: true)) {
                    return;
                }

                // The destination appeared after the existence check. Replace it below.
            }

            try {
                ExecuteWithRetry(() => File.Replace(temporaryPath, fullTargetPath, destinationBackupFileName: null));
                return;
            } catch (PlatformNotSupportedException) when (allowNonAtomicReplacementFallback) {
                // Fall back to a backup-and-move commit on file systems without File.Replace.
            } catch (IOException) when (allowNonAtomicReplacementFallback) {
                // Some file systems reject File.Replace even though moves are supported.
            }

            ReplaceUsingBackup(temporaryPath, fullTargetPath);
        }

        /// <summary>
        /// Atomically commits an existing staging file only when the destination can be claimed.
        /// </summary>
        /// <returns><c>false</c> when another writer already claimed the destination.</returns>
        public static bool TryCommitTemporaryFileIfAbsent(
            string temporaryPath,
            string targetPath) {
            if (string.IsNullOrWhiteSpace(temporaryPath)) {
                throw new ArgumentException("Temporary path cannot be empty.", nameof(temporaryPath));
            }

            EnsureTargetDirectory(targetPath);
            return TryMoveIfAbsent(temporaryPath, GetFullTargetPath(targetPath));
        }

        private static void EnsureDestinationWritable(string targetPath) {
            if (File.Exists(targetPath) && new FileInfo(targetPath).IsReadOnly) {
                throw new UnauthorizedAccessException($"Destination file '{targetPath}' is read-only.");
            }
        }

        /// <summary>Deletes a temporary file when it exists without hiding an earlier failure.</summary>
        public static void DeleteIfExists(string? path) {
            if (string.IsNullOrWhiteSpace(path)) return;

            try {
                if (File.Exists(path)) File.Delete(path);
            } catch {
                // Cleanup is best effort and must not hide the original save failure.
            }
        }

        private static string GetFullTargetPath(string targetPath) {
            if (string.IsNullOrWhiteSpace(targetPath)) throw new ArgumentException("Target path cannot be empty.", nameof(targetPath));

            string fullTargetPath = Path.GetFullPath(targetPath);
            if (string.IsNullOrEmpty(Path.GetFileName(fullTargetPath))) {
                throw new ArgumentException("Target path must include a file name.", nameof(targetPath));
            }

            return fullTargetPath;
        }

        /// <summary>Ensures the parent directory for a target file exists.</summary>
        public static void EnsureTargetDirectory(string targetPath) {
            string fullTargetPath = GetFullTargetPath(targetPath);
            string? directory = Path.GetDirectoryName(fullTargetPath);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        }

        private static string CreateBackupPath(string targetPath) {
            string? directory = Path.GetDirectoryName(targetPath);
            if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
            return Path.Combine(directory, $".{Path.GetFileName(targetPath)}.{Guid.NewGuid():N}.bak");
        }

        private static void ReplaceUsingBackup(string temporaryPath, string targetPath) {
            string backupPath = CreateBackupPath(targetPath);
            bool targetMoved = false;
            try {
                ExecuteWithRetry(() => File.Move(targetPath, backupPath));
                targetMoved = true;
                ExecuteWithRetry(() => File.Move(temporaryPath, targetPath));
                targetMoved = false;
                DeleteIfExists(backupPath);
            } catch (Exception commitException) {
                if (targetMoved && !File.Exists(targetPath) && File.Exists(backupPath)) {
                    try {
                        File.Move(backupPath, targetPath);
                        targetMoved = false;
                    } catch (Exception rollbackException) {
                        throw new IOException(
                            $"The new Office file could not be committed and the original destination could not be restored automatically. The original backup was retained at '{backupPath}'.",
                            new AggregateException(commitException, rollbackException));
                    }
                }

                throw;
            } finally {
                if (!targetMoved) DeleteIfExists(backupPath);
            }
        }

        private static void ExecuteWithRetry(Action operation) {
            for (int attempt = 0; ; attempt++) {
                try {
                    operation();
                    return;
                } catch (IOException) when (attempt < 9) {
                    Thread.Sleep(Math.Min(50 * (attempt + 1), 500));
                }
            }
        }

        private static bool TryMoveIfAbsent(
            string sourcePath,
            string targetPath,
            bool waitForClaim = false) {
            string claimPath = CreateClaimPath(targetPath);
            FileStream? claim = null;
            for (int attempt = 0; ; attempt++) {
                try {
                    claim = new FileStream(
                        claimPath,
                        FileMode.CreateNew,
                        FileAccess.ReadWrite,
                        FileShare.None);
                    break;
                } catch (IOException) when (File.Exists(claimPath) && !waitForClaim) {
                    return false;
                } catch (IOException) when (attempt < 9) {
                    Thread.Sleep(Math.Min(50 * (attempt + 1), 500));
                }
            }

            try {
                if (File.Exists(targetPath)) return false;
                try {
                    ExecuteWithRetry(() => File.Move(sourcePath, targetPath));
                    return true;
                } catch (IOException) when (File.Exists(targetPath)) {
                    return false;
                }
            } finally {
                claim.Dispose();
                DeleteIfExists(claimPath);
            }
        }

        private static string CreateClaimPath(string targetPath) {
            string? directory = Path.GetDirectoryName(targetPath);
            if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
            return Path.Combine(
                directory,
                $".{Path.GetFileName(targetPath)}.officeimo-commit");
        }
    }
}
