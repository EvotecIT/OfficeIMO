using System;

namespace OfficeIMO.Drawing;

/// <summary>Normalizes qualified legacy symbol-font list markers for Unicode text renderers.</summary>
internal static class OfficeTextListMarkerNormalizer {
    internal static (string Marker, bool UseTextFont) Normalize(string marker, string? fontFamily) {
        string trimmed = marker.Trim();
        if (trimmed.Length == 0) return (string.Empty, false);
        if (string.Equals(fontFamily, "Symbol", StringComparison.OrdinalIgnoreCase)) {
            return trimmed switch {
                "\uf0b7" or "\u00b7" => ("•", true),
                _ => (marker, false)
            };
        }
        if (string.Equals(fontFamily, "Wingdings", StringComparison.OrdinalIgnoreCase)) {
            return trimmed switch {
                "\uf0a7" => ("▪", true),
                _ => (marker, false)
            };
        }
        return (marker, false);
    }
}
