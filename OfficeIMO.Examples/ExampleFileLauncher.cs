using System.Diagnostics;

namespace OfficeIMO.Examples;

/// <summary>Example-only shell launcher for generated artifacts.</summary>
internal static class ExampleFileLauncher {
    internal static void Open(string filePath) {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        string fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) {
            throw new FileNotFoundException("File not found.", fullPath);
        }

        Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
    }
}
