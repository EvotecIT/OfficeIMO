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
            case "bolder": weight = inherited < 350 ? 400 : inherited < 550 ? 700 : Math.Max(900, inherited); return true;
            case "lighter": weight = inherited < 550 ? Math.Min(100, inherited) : inherited < 750 ? 400 : 700; return true;
        }
        // CSS weights are numbers from 1 to 1000; faces are matched on whole weights.
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            || !(number >= 1D && number <= 1000D)) return false; // also rejects NaN
        weight = (int)Math.Round(number, MidpointRounding.AwayFromZero);
        return true;
    }

    private static readonly string[] StretchNames = { "ultra-condensed", "extra-condensed", "condensed", "semi-condensed", "normal", "semi-expanded", "expanded", "extra-expanded", "ultra-expanded" };
    private static readonly double[] StretchWidths = { 50D, 62.5D, 75D, 87.5D, 100D, 112.5D, 125D, 150D, 200D };

    internal static bool TryStretch(string value, double inherited, out double stretch) {
        stretch = inherited;
        // Every resolved style asks; most carry no authored value.
        if (value.Length == 0) return true;
        string normalized = value.Trim().ToLowerInvariant();
        string[] names = StretchNames;
        double[] widths = StretchWidths;
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
        // A right-angle skew has no finite shear, so the bounds are exclusive; this also rejects NaN.
        return TryAngleDegrees(parts[1], out angle) && angle > -90D && angle < 90D;
    }

    // CSS <angle> in deg, grad, rad or turn, in degrees.
    private static bool TryAngleDegrees(string value, out double degrees) {
        degrees = 0D;
        (string Unit, double Scale)[] units = { ("deg", 1D), ("grad", 0.9D), ("rad", 180D / Math.PI), ("turn", 360D) };
        foreach ((string unit, double scale) in units) {
            if (!value.EndsWith(unit, StringComparison.Ordinal)) continue;
            if (!double.TryParse(value.Substring(0, value.Length - unit.Length), NumberStyles.Float, CultureInfo.InvariantCulture,
                    out double number)) return false;
            degrees = number * scale;
            return true;
        }
        return false;
    }
}
