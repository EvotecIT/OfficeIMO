using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>Launches saved Office artifacts through the operating system's registered application.</summary>
    internal static class OfficeFileLauncher {
        /// <summary>Opens an existing file through the operating system shell.</summary>
        public static void Open(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            }
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) {
                throw new FileNotFoundException("File not found.", fullPath);
            }

            Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
        }
    }
}
