using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeFileCommit {
        private static bool TryMoveIfAbsent(
            string sourcePath,
            string targetPath,
            bool waitForClaim = false,
            bool applyDefaultUnixPermissions = true) {
            string claimPath = CreateClaimPath(targetPath);
            bool windowsClaim = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
            FileStream? claim = null;
            for (int attempt = 0; ; attempt++) {
                try {
                    claim = OpenCommitClaim(claimPath, windowsClaim);
                    break;
                } catch (Exception exception) when (IsExistingClaimContention(exception, claimPath)) {
                    if (windowsClaim && TryDeleteAbandonedWindowsClaim(claimPath)) continue;
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
                if (applyDefaultUnixPermissions) OfficeTemporaryFile.ApplyDefaultUnixCreationMode(sourcePath);
                try {
                    ExecuteWithRetry(() => File.Move(sourcePath, targetPath));
                    return true;
                } catch (IOException) when (File.Exists(targetPath)) {
                    return false;
                }
            } finally {
                claim.Dispose();
                if (windowsClaim) DeleteIfExists(claimPath);
            }
        }

        private static FileStream OpenCommitClaim(string claimPath, bool windowsClaim) {
            if (windowsClaim) {
                // Preserve Windows kernel sharing and pathname deletion semantics. In
                // particular, DeleteOnClose would delete a claim symlink's target there.
                return new FileStream(claimPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            }

            // The Unix runtime checks the opened inode against the pathname, then unlinks
            // before releasing its exclusive lock. Separate reclamation or post-disposal
            // deletion can instead remove a successor's claim.
            try {
                return new FileStream(claimPath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                    FileShare.None, 4096, FileOptions.DeleteOnClose);
            } catch (UnauthorizedAccessException) {
                // A readable stale claim need not be writable when its parent permits
                // deletion. Prefer write access above for filesystems whose locks need it.
                return new FileStream(claimPath, FileMode.OpenOrCreate, FileAccess.Read,
                    FileShare.None, 4096, FileOptions.DeleteOnClose);
            }
        }

        private static bool IsExistingClaimContention(Exception exception, string claimPath) =>
            (exception is IOException || exception is UnauthorizedAccessException) &&
            File.Exists(claimPath);

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
        internal static string CreateClaimPath(string targetPath) {
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
