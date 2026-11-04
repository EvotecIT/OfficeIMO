using System.Runtime.InteropServices;

namespace OfficeIMO.Word;

/// <summary>
/// Resolves generic font family names to platform specific fonts.
/// </summary>
public static class WordFontResolver {
    // Generic families become document fonts that exist in Microsoft Office on every platform and have
    // metric-compatible or classified open substitutes, so the written name never depends on the
    // machine that generated the document.
    private static readonly Dictionary<string, string> _genericFonts = new(StringComparer.OrdinalIgnoreCase) {
        { "serif", "Times New Roman" },
        { "sans-serif", "Calibri" },
        { "monospace", "Consolas" },
        { "cursive", "Comic Sans MS" },
        { "fantasy", "Impact" }
    };

    private static readonly string[] _windowsFallbackFonts = {
        "Calibri",
        "Arial",
        "Times New Roman",
        "Consolas"
    };

    private static readonly string[] _linuxFallbackFonts = {
        "DejaVu Sans",
        "DejaVu Serif",
        "DejaVu Sans Mono"
    };

    private static readonly string[] _macFallbackFonts = {
        "Helvetica",
        "Arial",
        "Times",
        "Menlo"
    };

    /// <summary>
    /// Resolves the provided font family name to an actual installed font.
    /// Generic families like <c>serif</c> or <c>monospace</c> are mapped to
    /// portable Office families (Times New Roman, Calibri, Consolas, Comic Sans MS, Impact)
    /// that are the same on every platform.
    /// </summary>
    /// <param name="fontFamily">The requested font family.</param>
    /// <returns>The resolved font family or <paramref name="fontFamily"/> if no mapping exists.</returns>
    public static string? Resolve(string? fontFamily) {
        if (string.IsNullOrWhiteSpace(fontFamily)) {
            return null;
        }

        if (_genericFonts.TryGetValue(fontFamily!.Trim(), out string? generic)) {
            return generic;
        }

        // The portable generic targets are valid document names on every platform, installed or not.
        foreach (string portable in _genericFonts.Values) {
            if (string.Equals(portable, fontFamily.Trim(), StringComparison.OrdinalIgnoreCase)) return portable;
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            return ResolvePlatform(fontFamily!, _windowsFallbackFonts);
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
            return ResolvePlatform(fontFamily!, _linuxFallbackFonts);
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            return ResolvePlatform(fontFamily!, _macFallbackFonts);
        }

        return fontFamily;
    }

    private static string ResolvePlatform(string fontFamily, IEnumerable<string> fallbackFonts) {
        bool installed = IsFontInstalled(fontFamily);
        if (installed) {
            return fontFamily;
        }

        foreach (string fallback in fallbackFonts) {
            if (IsFontInstalled(fallback)) {
                return fallback;
            }
        }

        return fallbackFonts.FirstOrDefault() ?? fontFamily;
    }

    private static bool IsFontInstalled(string fontFamily) {
        try {
            return GetFontFiles().Any(file =>
                Path.GetFileNameWithoutExtension(file)
                    .IndexOf(fontFamily, StringComparison.OrdinalIgnoreCase) >= 0);
        } catch {
            return false;
        }
    }

    private static IEnumerable<string> GetFontFiles() {
        IEnumerable<string> paths;

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            string fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            paths = !string.IsNullOrEmpty(fontsDir) && Directory.Exists(fontsDir) ? new[] { fontsDir } : Array.Empty<string>();
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
            paths = new[] {
                "/usr/share/fonts",
                "/usr/local/share/fonts",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts")
            }.Where(Directory.Exists);
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            paths = new[] {
                "/System/Library/Fonts",
                "/Library/Fonts",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library/Fonts")
            }.Where(Directory.Exists);
        } else {
            paths = Array.Empty<string>();
        }

        string[] extensions = { "*.ttf", "*.ttc", "*.otf", "*.otc" };
        return paths.SelectMany(p =>
            extensions.SelectMany(ext =>
                Directory.EnumerateFiles(p, ext, SearchOption.AllDirectories)));
    }
}

