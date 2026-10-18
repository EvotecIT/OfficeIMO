using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.IO {
#else
namespace OfficeIMO.Core.Internal.IO {
#endif
    /// <summary>Owns the exclusive coordination-file lifetime for completed file publication.</summary>
    internal sealed class OfficeFileCommitClaim : IDisposable {
        private FileStream? _stream;
        private readonly string? _windowsClaimPath;

        private OfficeFileCommitClaim(FileStream stream, string? windowsClaimPath) {
            _stream = stream;
            _windowsClaimPath = windowsClaimPath;
        }

        /// <summary>Acquires a destination claim without waiting for a live owner.</summary>
        internal static OfficeFileCommitClaim Acquire(string destinationPath) {
            string claimPath = CreatePath(destinationPath);
            bool windows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
            while (true) {
                try {
                    FileStream stream = windows
                        // Preserve Windows sharing and pathname-deletion semantics. DeleteOnClose
                        // would instead follow a claim symlink and delete its target on Windows.
                        ? new FileStream(claimPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)
                        : OpenUnixClaim(claimPath);
                    return new OfficeFileCommitClaim(stream, windows ? claimPath : null);
                } catch (Exception exception) when (windows && IsExistingClaimContention(exception, claimPath)
                    && TryDeleteAbandonedWindowsClaim(claimPath)) {
                    // The stale Windows claim was removed; retry the exclusive create.
                }
            }
        }

        /// <summary>Identifies a claim collision so the caller can apply its own retry policy.</summary>
        internal static bool IsContention(Exception exception, string destinationPath) =>
            IsExistingClaimContention(exception, CreatePath(destinationPath));

        private static bool IsExistingClaimContention(Exception exception, string claimPath) =>
            (exception is IOException || exception is UnauthorizedAccessException) && File.Exists(claimPath);

        private static FileStream OpenUnixClaim(string claimPath) {
            // The runtime checks the opened inode against the pathname and unlinks it before
            // releasing its exclusive lock. Separate reclamation or post-disposal deletion can
            // instead remove a successor's claim.
            try {
                return new FileStream(claimPath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                    FileShare.None, 4096, FileOptions.DeleteOnClose);
            } catch (UnauthorizedAccessException) {
                // A readable stale claim need not be writable when its parent permits deletion.
                // Prefer write access above for filesystems whose locks require it.
                return new FileStream(claimPath, FileMode.OpenOrCreate, FileAccess.Read,
                    FileShare.None, 4096, FileOptions.DeleteOnClose);
            }
        }

        /// <summary>Releases a claim once, retaining the platform's established deletion behavior.</summary>
        public void Dispose() {
            FileStream? stream = Interlocked.Exchange(ref _stream, null);
            if (stream == null) return;
            stream.Dispose();
            if (_windowsClaimPath != null) {
                try {
                    if (File.Exists(_windowsClaimPath)) File.Delete(_windowsClaimPath);
                } catch {
                    // Best-effort cleanup must not hide the publication failure.
                }
            }
        }

        private static bool TryDeleteAbandonedWindowsClaim(string claimPath) {
            try {
                using (var abandonedClaim = new FileStream(
                           claimPath, FileMode.Open, FileAccess.Read, FileShare.Delete)) {
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

        /// <summary>Gets the exact coordination-file path used when committing a destination.</summary>
        internal static string CreatePath(string targetPath) {
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
