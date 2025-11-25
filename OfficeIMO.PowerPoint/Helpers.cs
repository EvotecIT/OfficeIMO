using System;
using System.Diagnostics;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Helper utilities for PowerPoint examples and library consumers.
    /// </summary>
    public static partial class Helpers {
        /// <summary>
        /// Opens the specified file with the OS default application when requested.
        /// </summary>
        /// <param name="filePath">Absolute or relative path to the file.</param>
        /// <param name="open">When <c>true</c>, launches the associated app.</param>
        public static void Open(string filePath, bool open) {
            if (!open) return;

            if (string.IsNullOrEmpty(filePath)) {
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
            }

            ProcessStartInfo startInfo = new ProcessStartInfo(filePath) {
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }
    }
}
