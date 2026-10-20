using System;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal.IO;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeFileCommit {
        private static bool TryMoveIfAbsent(
            string sourcePath,
            string targetPath,
            bool waitForClaim = false,
            bool applyDefaultUnixPermissions = true) {
            OfficeFileCommitClaim? claim = null;
            for (int attempt = 0; ; attempt++) {
                try {
                    claim = OfficeFileCommitClaim.Acquire(targetPath);
                    break;
                } catch (Exception exception) when (OfficeFileCommitClaim.IsContention(exception, targetPath)) {
                    if (!waitForClaim) return false;
                    if (attempt >= 9) throw;
                    Thread.Sleep(Math.Min(50 * (attempt + 1), 500));
                } catch (IOException) when (attempt < 9) {
                    Thread.Sleep(Math.Min(50 * (attempt + 1), 500));
                }
            }

            using (claim) {
                if (File.Exists(targetPath)) return false;
                if (applyDefaultUnixPermissions) OfficeTemporaryFile.ApplyDefaultUnixCreationMode(sourcePath);
                try {
                    ExecuteWithRetry(() => File.Move(sourcePath, targetPath));
                    return true;
                } catch (IOException) when (File.Exists(targetPath)) {
                    return false;
                }
            }
        }
    }
}
