using System.Diagnostics;

namespace OfficeIMO.Examples;

/// <summary>Example-only shell launcher for generated artifacts.</summary>
internal static class ExampleFileLauncher {
    internal static void Open(string filePath) {
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
