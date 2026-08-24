using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
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
            string temporaryPath = string.Empty;
            try {
                using (var stream = CreateTemporaryFile(targetPath, FileOptions.None, out temporaryPath)) {
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
            string temporaryPath = string.Empty;
            try {
                using (var stream = CreateTemporaryFile(fullTargetPath, FileOptions.None, out temporaryPath)) {
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
            string temporaryPath = string.Empty;
            try {
                using (var stream = CreateTemporaryFile(targetPath, FileOptions.Asynchronous, out temporaryPath, 8192)) {
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
            string temporaryPath = string.Empty;
            try {
                using (var stream = CreateTemporaryFile(fullTargetPath, FileOptions.Asynchronous, out temporaryPath, 8192)) {
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
            string temporaryPath = string.Empty;
            try {
                using (var stream = CreateTemporaryFile(targetPath, FileOptions.Asynchronous, out temporaryPath, 8192)) {
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
            string temporaryPath = string.Empty;
            try {
                using (var stream = CreateTemporaryFile(targetPath, FileOptions.Asynchronous, out temporaryPath, 8192)) {
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

            return Path.Combine(directory, $".officeimo-{Guid.NewGuid():N}.tmp");
        }

        /// <summary>Creates an owner-only same-directory staging file suitable for an atomic commit.</summary>
        public static FileStream CreateTemporaryFile(
            string targetPath,
            FileOptions options,
            out string temporaryPath,
            int bufferSize = 81920) {
            EnsureTargetDirectory(targetPath);
            temporaryPath = CreateTemporaryPath(targetPath);
            return OfficeTemporaryFile.CreateAtPath(temporaryPath, bufferSize, options);
        }

        /// <summary>Creates a same-directory staging path that preserves the destination extension.</summary>
        public static string CreateStagingPath(string targetPath) {
            string fullTargetPath = GetFullTargetPath(targetPath);
            string? directory = Path.GetDirectoryName(fullTargetPath);
            if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();

            string extension = Path.GetExtension(fullTargetPath);
            return Path.Combine(directory, $".officeimo-{Guid.NewGuid():N}{extension}");
        }

        /// <summary>Commits a completed temporary file to its destination.</summary>
        public static void CommitTemporaryFile(
            string temporaryPath,
            string targetPath,
            ConflictPolicy conflictPolicy = ConflictPolicy.Replace) {
            CommitTemporaryFileCore(temporaryPath, targetPath, conflictPolicy,
                allowNonAtomicReplacementFallback: true,
                allowReadOnlyUnixDestination: false);
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
                allowNonAtomicReplacementFallback: false,
                allowReadOnlyUnixDestination: false);
        }

        /// <summary>
        /// Atomically replaces a Unix destination whose inode is read-only when the caller owns a
        /// separate mutation lock and has already verified that the parent directory is writable.
        /// Windows destinations retain the ordinary read-only guard.
        /// </summary>
        internal static void CommitTemporaryFileAtomicallyForLockedMutation(
            string temporaryPath,
            string targetPath) {
            CommitTemporaryFileCore(temporaryPath, targetPath, ConflictPolicy.Replace,
                allowNonAtomicReplacementFallback: false,
                allowReadOnlyUnixDestination: true);
        }

        /// <summary>
        /// Atomically installs a completed staging file only when the displaced destination still
        /// matches the caller's expected snapshot and, when requested, the installed file matches
        /// the validated staging snapshot. A mismatch restores the displaced destination.
        /// </summary>
        public static bool TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
            string temporaryPath,
            string targetPath,
            Func<string, bool> destinationMatchesExpected,
            Func<string, bool>? installedFileMatchesExpected = null) =>
            TryCommitTemporaryFileAtomicallyIfDestinationUnchangedCore(
                temporaryPath,
                targetPath,
                destinationMatchesExpected,
                installedFileMatchesExpected,
                afterFirstRollbackReplacement: null);

        internal static bool TryCommitTemporaryFileAtomicallyIfDestinationUnchangedForTesting(
            string temporaryPath,
            string targetPath,
            Func<string, bool> destinationMatchesExpected,
            Func<string, bool>? installedFileMatchesExpected,
            Action<string> afterFirstRollbackReplacement) =>
            TryCommitTemporaryFileAtomicallyIfDestinationUnchangedCore(
                temporaryPath,
                targetPath,
                destinationMatchesExpected,
                installedFileMatchesExpected,
                afterFirstRollbackReplacement);

        private static bool TryCommitTemporaryFileAtomicallyIfDestinationUnchangedCore(
            string temporaryPath,
            string targetPath,
            Func<string, bool> destinationMatchesExpected,
            Func<string, bool>? installedFileMatchesExpected,
            Action<string>? afterFirstRollbackReplacement) {
            if (string.IsNullOrWhiteSpace(temporaryPath)) {
                throw new ArgumentException("Temporary path cannot be empty.", nameof(temporaryPath));
            }
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(destinationMatchesExpected);
#else
            if (destinationMatchesExpected == null) throw new ArgumentNullException(nameof(destinationMatchesExpected));
#endif

            string fullTargetPath = GetFullTargetPath(targetPath);
            EnsureDestinationWritable(fullTargetPath);
            string backupPath = CreateTemporaryPath(fullTargetPath);
            string displacedPath = CreateTemporaryPath(fullTargetPath);
            bool targetContainsTemporary = false;
            bool preserveBackupPath = false;
            bool preserveDisplacedPath = false;
            string installedTemporaryIdentity = ComputeFileIdentity(temporaryPath);
            try {
#if NET6_0_OR_GREATER
                if (!OperatingSystem.IsWindows()) {
                    File.SetUnixFileMode(temporaryPath, File.GetUnixFileMode(fullTargetPath));
                }
#endif
                ExecuteWithRetry(() => File.Replace(temporaryPath, fullTargetPath, backupPath));
                targetContainsTemporary = true;
                if (destinationMatchesExpected(backupPath) &&
                    (installedFileMatchesExpected == null || installedFileMatchesExpected(fullTargetPath))) {
                    DeleteIfExists(backupPath);
                    targetContainsTemporary = false;
                    return true;
                }

                RestoreDisplacedDestinationWithoutLosingConcurrentSave(
                    fullTargetPath,
                    backupPath,
                    displacedPath,
                    installedTemporaryIdentity,
                    ref targetContainsTemporary,
                    ref preserveBackupPath,
                    afterFirstRollbackReplacement,
                    ref preserveDisplacedPath);
                return false;
            } catch (Exception commitException) {
                if (targetContainsTemporary && File.Exists(backupPath) && File.Exists(fullTargetPath)) {
                    try {
                        RestoreDisplacedDestinationWithoutLosingConcurrentSave(
                            fullTargetPath,
                            backupPath,
                            displacedPath,
                            installedTemporaryIdentity,
                            ref targetContainsTemporary,
                            ref preserveBackupPath,
                            afterFirstRollbackReplacement,
                            ref preserveDisplacedPath);
                    } catch (Exception rollbackException) {
                        throw new IOException(
                            "The guarded atomic commit failed and the displaced destination could not be restored. " +
                            "Its recoverable files remain at '" + backupPath + "' and '" + displacedPath + "'.",
                            new AggregateException(commitException, rollbackException));
                    }
                }
                throw;
            } finally {
                if (!targetContainsTemporary && !preserveBackupPath) DeleteIfExists(backupPath);
                if (!preserveDisplacedPath) DeleteIfExists(displacedPath);
            }
        }

        private static void RestoreDisplacedDestinationWithoutLosingConcurrentSave(
            string targetPath,
            string backupPath,
            string displacedPath,
            string installedTemporaryIdentity,
            ref bool targetContainsTemporary,
            ref bool preserveBackupPath,
            Action<string>? afterFirstRollbackReplacement,
            ref bool preserveDisplacedPath) {
            string restoredDestinationIdentity = ComputeFileIdentity(backupPath);
            ExecuteWithRetry(() => File.Replace(backupPath, targetPath, displacedPath));
            targetContainsTemporary = false;
            afterFirstRollbackReplacement?.Invoke(targetPath);
            if (installedTemporaryIdentity.Length == 0 ||
                string.Equals(installedTemporaryIdentity, ComputeFileIdentity(displacedPath), StringComparison.Ordinal)) {
                DeleteIfExists(displacedPath);
                return;
            }

            // A different writer replaced or rewrote the target while the caller was checking the
            // displaced destination. Put that newer save back instead of silently overwriting it.
            preserveDisplacedPath = true;
            ExecuteWithRetry(() => File.Replace(displacedPath, targetPath, backupPath));
            preserveDisplacedPath = false;
            preserveBackupPath = true;
            if (string.Equals(restoredDestinationIdentity, ComputeFileIdentity(backupPath), StringComparison.Ordinal)) {
                preserveBackupPath = false;
                DeleteIfExists(backupPath);
                return;
            }

            throw new IOException(
                "A newer concurrent save was displaced during guarded rollback and remains recoverable at '" +
                backupPath + "'.");
        }

        private static string ComputeFileIdentity(string path) {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
            using SHA256 sha256 = SHA256.Create();
            return Convert.ToBase64String(sha256.ComputeHash(stream));
        }

        private static void CommitTemporaryFileCore(
            string temporaryPath,
            string targetPath,
            ConflictPolicy conflictPolicy,
            bool allowNonAtomicReplacementFallback,
            bool allowReadOnlyUnixDestination) {
            if (string.IsNullOrWhiteSpace(temporaryPath)) throw new ArgumentException("Temporary path cannot be empty.", nameof(temporaryPath));

            string fullTargetPath = GetFullTargetPath(targetPath);
            if (conflictPolicy == ConflictPolicy.FailIfExists) {
                if (!TryMoveIfAbsent(temporaryPath, fullTargetPath)) {
                    throw new IOException($"Destination file '{fullTargetPath}' already exists.");
                }
                return;
            }

            if (!allowReadOnlyUnixDestination || RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                EnsureDestinationWritable(fullTargetPath);
            }

            if (!File.Exists(fullTargetPath)) {
                if (TryMoveIfAbsent(temporaryPath, fullTargetPath, waitForClaim: true)) {
                    return;
                }

                // The destination appeared after the existence check. Replace it below.
            }

            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                // A rename/replace installs the staging inode on Unix. Apply the existing
                // destination's mode first so a restrictive workbook or document cannot be
                // widened to the staging file's default umask-derived permissions.
                OfficeTemporaryFile.CopyUnixFileMode(fullTargetPath, temporaryPath);
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
            return Path.Combine(directory, $".officeimo-{Guid.NewGuid():N}.bak");
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
                } catch (Exception exception) when (IsExistingClaimContention(exception, claimPath)) {
                    if (TryDeleteAbandonedClaim(claimPath)) {
                        continue;
                    }
                    if (!waitForClaim) return false;
                    if (attempt >= 9) throw;
                    Thread.Sleep(Math.Min(50 * (attempt + 1), 500));
                } catch (UnauthorizedAccessException) when (attempt < 9) {
                    // Windows can report a sharing violation as access denied while
                    // another writer owns the claim, and File.Exists may transiently
                    // return false for that locked path. Retry so the owner can release
                    // the claim; a persistent permission failure is rethrown below.
                    Thread.Sleep(Math.Min(50 * (attempt + 1), 500));
                } catch (IOException) when (attempt < 9) {
                    Thread.Sleep(Math.Min(50 * (attempt + 1), 500));
                }
            }

            try {
                if (File.Exists(targetPath)) return false;
                OfficeTemporaryFile.ApplyDefaultUnixCreationMode(sourcePath);
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

        private static bool IsExistingClaimContention(Exception exception, string claimPath) =>
            (exception is IOException || exception is UnauthorizedAccessException) &&
            File.Exists(claimPath);

        private static bool TryDeleteAbandonedClaim(string claimPath) {
            try {
                // A live committer holds the claim with FileShare.None, so this open
                // succeeds only after that owner exits or crashes.
                using (var abandonedClaim = new FileStream(
                           claimPath,
                           FileMode.Open,
                           FileAccess.Read,
                           FileShare.Delete)) {
                    File.Delete(claimPath);
                }
                return true;
            } catch (FileNotFoundException) {
                return true;
            } catch (DirectoryNotFoundException) {
                return true;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (IOException) {
                return false;
            }
        }

        private static string CreateClaimPath(string targetPath) {
            string? directory = Path.GetDirectoryName(targetPath);
            if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
            string legacyClaimName = "." + Path.GetFileName(targetPath) + ".officeimo-commit";
            string legacyClaimPath = Path.Combine(directory, legacyClaimName);
            if (Encoding.UTF8.GetByteCount(legacyClaimName) <= 255 &&
                (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) || legacyClaimPath.Length < 260)) {
                // Preserve coordination with OfficeIMO versions that use the destination-derived
                // claim name. Fall back to the compact hash only when that legacy path is not
                // portable enough to create.
                return legacyClaimPath;
            }

            string fullTargetPath = Path.GetFullPath(targetPath);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                fullTargetPath = fullTargetPath.ToUpperInvariant();
            }

            byte[] pathBytes = Encoding.UTF8.GetBytes(fullTargetPath);
            byte[] hash;
            using (SHA256 sha256 = SHA256.Create()) {
                hash = sha256.ComputeHash(pathBytes);
            }

            const string hex = "0123456789abcdef";
            var token = new char[24];
            for (int index = 0; index < token.Length / 2; index++) {
                token[index * 2] = hex[hash[index] >> 4];
                token[(index * 2) + 1] = hex[hash[index] & 0x0F];
            }

            return Path.Combine(directory, ".officeimo-" + new string(token) + ".commit");
        }
    }
}
