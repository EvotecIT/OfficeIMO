using System;
using System.IO;
using OfficeIMO.Reader.Internal.IO;

namespace OfficeIMO.Reader;

/// <summary>
/// Stages Reader materializations and coordinates publication through the shared destination claim.
/// </summary>
internal static class ReaderFileCommit {
    /// <summary>Publishes completed bytes, preserving existing destinations when overwrite is disabled.</summary>
        internal static bool WriteAllBytes(string path, byte[] bytes, bool overwrite = true) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));

        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(directory)) {
            throw new ArgumentException("The output path must include a directory.", nameof(path));
        }

        Directory.CreateDirectory(directory);
        string temporaryPath = Path.Combine(
            directory,
            "." + Path.GetFileName(fullPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            File.WriteAllBytes(temporaryPath, bytes);
            if (!overwrite) {
                if (!OfficeFileCommitClaim.TryAcquire(fullPath, out OfficeFileCommitClaim? claim)) return false;
                using (claim) {
                    if (File.Exists(fullPath)) {
                        return false;
                    }
                    try { File.Move(temporaryPath, fullPath); }
                    catch (IOException) when (File.Exists(fullPath)) { return false; }
                }
            } else if (File.Exists(fullPath)) {
                File.Replace(temporaryPath, fullPath, null);
            } else {
                File.Move(temporaryPath, fullPath);
            }
            return true;
        } finally {
            try {
                if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
            } catch {
                // Best-effort cleanup must not hide the publication failure.
            }
        }
    }
}
