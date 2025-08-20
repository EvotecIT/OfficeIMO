using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;

namespace OfficeIMO.Word;

/// <summary>
/// Resolves generic font family names to platform specific fonts.
/// </summary>
public static class FontResolver {
    private static readonly Dictionary<string, string> _windowsFonts = new(StringComparer.OrdinalIgnoreCase) {
        { "serif", "Times New Roman" },
        { "sans-serif", "Calibri" },
        { "monospace", "Consolas" },
        { "cursive", "Comic Sans MS" },
        { "fantasy", "Impact" }
    };

    private static readonly Dictionary<string, string> _linuxFonts = new(StringComparer.OrdinalIgnoreCase) {
        { "serif", "DejaVu Serif" },
        { "sans-serif", "DejaVu Sans" },
        { "monospace", "DejaVu Sans Mono" },
        { "cursive", "DejaVu Sans" },
        { "fantasy", "DejaVu Sans" }
    };

    private static readonly Dictionary<string, string> _macFonts = new(StringComparer.OrdinalIgnoreCase) {
        { "serif", "Times" },
        { "sans-serif", "Helvetica" },
        { "monospace", "Menlo" },
        { "cursive", "Apple Chancery" },
        { "fantasy", "Papyrus" }
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
    /// platform specific fonts.
    /// </summary>
    /// <param name="fontFamily">The requested font family.</param>
    /// <returns>The resolved font family or <paramref name="fontFamily"/> if no mapping exists.</returns>
    public static string? Resolve(string? fontFamily) {
        if (string.IsNullOrWhiteSpace(fontFamily)) {
            return null;
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            return ResolvePlatform(fontFamily, _windowsFonts, _windowsFallbackFonts);
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
            return ResolvePlatform(fontFamily, _linuxFonts, _linuxFallbackFonts);
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            return ResolvePlatform(fontFamily, _macFonts, _macFallbackFonts);
        }

        return fontFamily;
    }

    private static string ResolvePlatform(string fontFamily, Dictionary<string, string> genericFonts, IEnumerable<string> fallbackFonts) {
        if (genericFonts.TryGetValue(fontFamily, out string value)) {
            fontFamily = value;
        }

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
                    .Contains(fontFamily, StringComparison.OrdinalIgnoreCase));
        } catch {
            return false;
        }
    }

    private static IEnumerable<string> GetFontFiles() {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            string? fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            if (Directory.Exists(fontsDir)) {
                return Directory.EnumerateFiles(fontsDir, "*.ttf", SearchOption.TopDirectoryOnly);
            }
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
            var paths = new[] {
                "/usr/share/fonts",
                "/usr/local/share/fonts",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts")
            };
            return paths.Where(Directory.Exists)
                .SelectMany(p => Directory.EnumerateFiles(p, "*.ttf", SearchOption.AllDirectories));
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            var paths = new[] {
                "/System/Library/Fonts",
                "/Library/Fonts",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library/Fonts")
            };
            return paths.Where(Directory.Exists)
                .SelectMany(p => Directory.EnumerateFiles(p, "*.ttf", SearchOption.AllDirectories));
        }

        return Array.Empty<string>();
    }
}

