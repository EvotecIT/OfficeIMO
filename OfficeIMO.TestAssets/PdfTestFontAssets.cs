using System;
using System.IO;

namespace OfficeIMO.TestAssets;

internal static class PdfTestFontAssets {
    private const string BundledFontFileName = "SourceSerif4-Regular.otf";

    internal static byte[] LoadBundledOpenTypeCffFont() {
        string copiedPath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Fonts", BundledFontFileName);
        if (File.Exists(copiedPath)) {
            return File.ReadAllBytes(copiedPath);
        }

        const string repositoryRelativePath = "OfficeIMO.Pdf.Tests/Pdf/Fixtures/Fonts/SourceSerif4-Regular.otf";
        foreach (string root in EnumerateSearchRoots()) {
            string candidate = Path.Combine(root, repositoryRelativePath.Replace('/', Path.DirectorySeparatorChar));
            if (File.Exists(candidate)) {
                return File.ReadAllBytes(candidate);
            }
        }

        throw new FileNotFoundException(
            "The bundled PDF test font could not be found in the test output or repository checkout.",
            BundledFontFileName);
    }

    private static System.Collections.Generic.IEnumerable<string> EnumerateSearchRoots() {
        string? current = AppContext.BaseDirectory;
        while (!string.IsNullOrWhiteSpace(current)) {
            yield return current;
            current = Directory.GetParent(current)?.FullName;
        }

        current = Directory.GetCurrentDirectory();
        while (!string.IsNullOrWhiteSpace(current)) {
            yield return current;
            current = Directory.GetParent(current)?.FullName;
        }
    }
}
