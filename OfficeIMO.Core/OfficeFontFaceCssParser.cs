using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>Shared CSS face requests used by SVG and HTML readers before font matching.</summary>
internal static class OfficeFontFaceCssParser {
    internal static bool TryWeight(string value, int inherited, out int weight) {
        string normalized = value.Trim().ToLowerInvariant();
        weight = inherited;
        switch (normalized) {
            case "": case "inherit": return true;
            case "normal": weight = 400; return true;
            case "bold": weight = 700; return true;
            case "bolder": weight = inherited < 350 ? 400 : inherited < 550 ? 700 : 900; return true;
            case "lighter": weight = inherited < 550 ? 100 : inherited < 750 ? 400 : 700; return true;
        }
        return int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out weight) && weight >= 1 && weight <= 1000;
    }

    internal static bool TryStretch(string value, double inherited, out double stretch) {
        string normalized = value.Trim().ToLowerInvariant();
        string[] names = { "ultra-condensed", "extra-condensed", "condensed", "semi-condensed", "normal", "semi-expanded", "expanded", "extra-expanded", "ultra-expanded" };
        double[] widths = { 50D, 62.5D, 75D, 87.5D, 100D, 112.5D, 125D, 150D, 200D };
        stretch = inherited;
        if (normalized.Length == 0 || normalized == "inherit") return true;
        for (int i = 0; i < names.Length; i++) if (normalized == names[i]) { stretch = widths[i]; return true; }
        if (normalized == "wider") {
            foreach (double width in widths) if (width > inherited + .0001D) { stretch = width; return true; }
            stretch = 200D; return true;
        }
        if (normalized == "narrower") {
            for (int i = widths.Length - 1; i >= 0; i--) if (widths[i] < inherited - .0001D) { stretch = widths[i]; return true; }
            stretch = 50D; return true;
        }
        return normalized.EndsWith("%", StringComparison.Ordinal) && double.TryParse(normalized.Substring(0, normalized.Length - 1),
            NumberStyles.Float, CultureInfo.InvariantCulture, out stretch) && stretch >= 50D && stretch <= 200D;
    }

    internal static bool TrySlant(string value, OfficeFontFaceDescriptor inherited, out OfficeFontSlant slant, out double angle) {
        string normalized = value.Trim().ToLowerInvariant();
        slant = inherited.Slant; angle = inherited.ObliqueAngleDegrees;
        if (normalized.Length == 0 || normalized == "inherit") return true;
        if (normalized == "normal") { slant = OfficeFontSlant.Normal; angle = 0D; return true; }
        if (normalized == "italic") { slant = OfficeFontSlant.Italic; angle = 0D; return true; }
        string[] parts = normalized.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0 || parts[0] != "oblique" || parts.Length > 2) return false;
        slant = OfficeFontSlant.Oblique; angle = 14D;
        if (parts.Length == 1) return true;
        return parts[1].EndsWith("deg", StringComparison.Ordinal) && double.TryParse(parts[1].Substring(0, parts[1].Length - 3),
            NumberStyles.Float, CultureInfo.InvariantCulture, out angle) && angle > -90D && angle < 90D;
    }
}
