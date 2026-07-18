using System;
using System.IO;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfComplianceTestFonts {
    internal static string? FindBundledOpenTypeCffFont() {
        const string relativePath = "OfficeIMO.Pdf.Tests/Pdf/Fixtures/Fonts/SourceSerif4-Regular.otf";
        foreach (string root in EnumerateSearchRoots()) {
            string candidate = Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar));
            if (File.Exists(candidate)) {
                return candidate;
            }
        }

        return null;
    }

    internal static string? FindLocalTrueTypeFont() {
        string windowsFont = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (File.Exists(windowsFont)) {
            return windowsFont;
        }

        string[] candidates = {
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf"
        };

        foreach (string candidate in candidates) {
            if (File.Exists(candidate)) {
                return candidate;
            }
        }

        return null;
    }

    private static System.Collections.Generic.IEnumerable<string> EnumerateSearchRoots() {
        string? current = AppContext.BaseDirectory;
        while (!string.IsNullOrWhiteSpace(current)) {
            yield return current;
            DirectoryInfo? parent = Directory.GetParent(current);
            current = parent?.FullName;
        }

        current = Directory.GetCurrentDirectory();
        while (!string.IsNullOrWhiteSpace(current)) {
            yield return current;
            DirectoryInfo? parent = Directory.GetParent(current);
            current = parent?.FullName;
        }
    }
}
