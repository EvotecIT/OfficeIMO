using System;
using System.IO;

namespace OfficeIMO.Reader;

/// <summary>
/// Core-owned atomic file publication used by Reader materializers.
/// </summary>
internal static class ReaderFileCommit {
    internal static void WriteAllBytes(string path, byte[] bytes) {
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
            if (File.Exists(fullPath)) {
                File.Delete(fullPath);
            }
            File.Move(temporaryPath, fullPath);
        } finally {
            try {
                if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
            } catch {
                // Best-effort cleanup must not hide the publication failure.
            }
        }
    }
}
