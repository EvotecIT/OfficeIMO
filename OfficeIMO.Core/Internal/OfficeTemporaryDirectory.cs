using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeIMO.Core.Internal {
    /// <summary>Creates task-owned temporary directories, with owner-only access on Unix.</summary>
    internal static class OfficeTemporaryDirectory {
        internal static string Create(string prefix, string? parentDirectory = null) {
            if (string.IsNullOrWhiteSpace(prefix) || prefix.IndexOfAny(new[] { '/', '\\', ':' }) >= 0) {
                throw new ArgumentException("A temporary directory prefix must be a simple name.", nameof(prefix));
            }
            string parent = parentDirectory is null ? Path.GetTempPath() : Path.GetFullPath(parentDirectory);
            if (parentDirectory is not null) Directory.CreateDirectory(parent);
            for (int attempt = 0; attempt < 16; attempt++) {
                string path = Path.Combine(parent, prefix + Guid.NewGuid().ToString("N"));
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                    Directory.CreateDirectory(path);
                    return path;
                }
                if (UnixMkdir(path, 0x1C0U) == 0) return path; // 0700, before any content is created.
                int error = Marshal.GetLastWin32Error();
                if (error != 17) throw new IOException($"Unable to create a private temporary directory (errno {error}).");
            }
            throw new IOException("Unable to allocate a unique private temporary directory.");
        }

        [DllImport("libc", SetLastError = true, EntryPoint = "mkdir")]
        private static extern int UnixMkdir(string path, uint mode);
    }
}
