using System;
using System.IO;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfComplianceTestFonts {
    internal static string? FindLocalTrueTypeFont() {
        string windowsFont = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (File.Exists(windowsFont)) {
            return windowsFont;
        }

        string[] candidates = {
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/Library/Fonts/Arial.ttf"
        };

        foreach (string candidate in candidates) {
            if (File.Exists(candidate)) {
                return candidate;
            }
        }

        return null;
    }
}
