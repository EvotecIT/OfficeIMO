using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OfficeIMO.Word;

/// <summary>
/// Resolves generic font family names to platform specific fonts.
/// </summary>
public static class FontResolver {
    private static readonly Dictionary<string, string> _windowsFonts = new(StringComparer.OrdinalIgnoreCase) {
        { "serif", "Times New Roman" },
        { "sans-serif", "Calibri" },
        { "monospace", "Consolas" }
    };

    private static readonly Dictionary<string, string> _linuxFonts = new(StringComparer.OrdinalIgnoreCase) {
        { "serif", "DejaVu Serif" },
        { "sans-serif", "DejaVu Sans" },
        { "monospace", "DejaVu Sans Mono" }
    };

    private static readonly Dictionary<string, string> _macFonts = new(StringComparer.OrdinalIgnoreCase) {
        { "serif", "Times" },
        { "sans-serif", "Helvetica" },
        { "monospace", "Menlo" }
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
            if (_windowsFonts.TryGetValue(fontFamily, out string value)) {
                return value;
            }
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
            if (_linuxFonts.TryGetValue(fontFamily, out string value)) {
                return value;
            }
        } else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            if (_macFonts.TryGetValue(fontFamily, out string value)) {
                return value;
            }
        }

        return fontFamily;
    }
}

