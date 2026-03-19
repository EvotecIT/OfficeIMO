using System.Runtime.InteropServices;

namespace OfficeIMO.MarkdownRenderer.Wpf;

/// <summary>
/// Exposes platform support helpers for the WPF markdown host package.
/// </summary>
public static class MarkdownRendererWpfPlatform {
    /// <summary>
    /// Returns <see langword="true"/> when the current process is running on Windows.
    /// </summary>
    public static bool IsSupported => RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
}
