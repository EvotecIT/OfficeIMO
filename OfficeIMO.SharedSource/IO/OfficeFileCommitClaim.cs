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
        private static readonly byte[] ClaimPrefix = Encoding.ASCII.GetBytes("OfficeIMO-commit-claim-v1:");
        private FileStream? _stream;
        private readonly string? _windowsClaimPath;
        private readonly byte[]? _windowsClaimMarker;

        private OfficeFileCommitClaim(FileStream stream, string? windowsClaimPath, byte[]? windowsClaimMarker) {
            _stream = stream;
            _windowsClaimPath = windowsClaimPath;
            _windowsClaimMarker = windowsClaimMarker;
        }

        /// <summary>Acquires a destination claim with bounded Windows deletion-pending retries.</summary>
        internal static OfficeFileCommitClaim Acquire(string destinationPath) {
            string claimPath = CreatePath(destinationPath);
            bool windows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
            const int maximumRetries = 9;
            for (int attempt = 0; ; attempt++) {
                try {
                    FileStream stream = windows
                        // Preserve Windows sharing and pathname-deletion semantics. DeleteOnClose
                        // would instead follow a claim symlink and delete its target on Windows.
                        ? new FileStream(claimPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)
                        : OpenUnixClaim(claimPath);
                    if (!windows) return new OfficeFileCommitClaim(stream, null, null);
                    try {
                        byte[] marker = new byte[ClaimPrefix.Length + 16];
                        Buffer.BlockCopy(ClaimPrefix, 0, marker, 0, ClaimPrefix.Length);
                        byte[] token = Guid.NewGuid().ToByteArray();
                        Buffer.BlockCopy(token, 0, marker, ClaimPrefix.Length, token.Length);
                        stream.Write(marker, 0, marker.Length);
                        stream.Flush(true);
                        return new OfficeFileCommitClaim(stream, claimPath, marker);
                    } catch {
                        stream.Dispose();
                        throw;
                    }
                } catch (Exception exception) when (windows && attempt < maximumRetries &&
                    IsExistingClaimContention(exception, claimPath) && TryDeleteAbandonedWindowsClaim(claimPath)) {
                    // A closed orphan was removed. Retain the exclusive create
                    // and the marker check so a new owner's claim cannot be erased.
                } catch (UnauthorizedAccessException exception) when (windows && attempt < maximumRetries &&
                    IsPossiblePendingDeletion(exception, claimPath)) {
                    // DeleteFile leaves the pathname pending until the disposing handle closes.
                    // During that interval CreateNew reports access denied and File.Exists may
                    // report false. Retry this ambiguous Windows error for a bounded interval;
                    // persistent ACL, read-only, and directory failures retain their exception.
                    Thread.Sleep(1);
                }
            }
        }

        private static bool TryDeleteAbandonedWindowsClaim(string claimPath) {
            try {
                if ((File.GetAttributes(claimPath) & (FileAttributes.ReparsePoint | FileAttributes.Directory)) != 0)
                    return false;
                using (var inactive = new FileStream(claimPath, FileMode.Open, FileAccess.Read, FileShare.Delete)) {
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

        /// <summary>Identifies a claim collision so the caller can apply its own retry policy.</summary>
        internal static bool IsContention(Exception exception, string destinationPath) =>
            IsExistingClaimContention(exception, CreatePath(destinationPath));

        private static bool IsExistingClaimContention(Exception exception, string claimPath) =>
            (exception is IOException || exception is UnauthorizedAccessException) && File.Exists(claimPath);

        private static bool IsPossiblePendingDeletion(UnauthorizedAccessException exception, string claimPath) {
            if (exception.HResult != unchecked((int)0x80070005) || Directory.Exists(claimPath)) return false;
            try {
                return (File.GetAttributes(claimPath) & (FileAttributes.ReadOnly | FileAttributes.Directory)) == 0;
            } catch (IOException) {
                return true;
            } catch (UnauthorizedAccessException) {
                // Attribute lookup also fails while deletion is pending. An ACL failure is
                // indistinguishable here, but retains its original exception after nine pauses.
                return true;
            }
        }

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
            if (_windowsClaimPath != null && _windowsClaimMarker != null) {
                try {
                    using (var ownedClaim = new FileStream(_windowsClaimPath, FileMode.Open, FileAccess.Read, FileShare.Delete)) {
                        if (HasClaimMarker(ownedClaim, _windowsClaimMarker)) File.Delete(_windowsClaimPath);
                    }
                } catch {
                    // Best-effort cleanup must not hide the publication failure.
                }
            }
        }

        private static bool HasClaimMarker(FileStream stream, byte[] expected) {
            if (stream.Length != expected.Length) return false;
            var actual = new byte[expected.Length];
            int read = 0;
            while (read < expected.Length) {
                int count = stream.Read(actual, read, expected.Length - read);
                if (count == 0) return false;
                read += count;
            }
            for (int i = 0; i < expected.Length; i++)
                if (actual[i] != expected[i]) return false;
            return true;
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
