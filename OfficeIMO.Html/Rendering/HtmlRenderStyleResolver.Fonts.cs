using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private static OfficeFontFaceDescriptor ResolveFontFaceDescriptor(
        string tag,
        HtmlComputedStyle computed,
        OfficeFontFaceDescriptor inherited) {
        int weight = ResolveRequestedFontWeight(tag, computed.GetValue("font-weight"), inherited.Weight);
        double stretch = ResolveRequestedFontStretch(computed.GetValue("font-stretch"), inherited.StretchPercent);
        ResolveRequestedFontSlant(
            tag,
            computed.GetValue("font-style"),
            inherited,
            out OfficeFontSlant slant,
            out double obliqueAngle);
        return new OfficeFontFaceDescriptor(weight, stretch, slant, obliqueAngle);
    }

    private static int ResolveRequestedFontWeight(string tag, string value, int inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0) {
            bool heading = tag.Length == 2 && tag[0] == 'h' && tag[1] >= '1' && tag[1] <= '6';
            return heading || tag == "b" || tag == "strong" ? 700 : inherited;
        }
        if (normalized == "normal") return 400;
        if (normalized == "bold") return 700;
        if (normalized == "bolder") {
            if (inherited < 350) return 400;
            if (inherited < 550) return 700;
            return 900;
        }
        if (normalized == "lighter") {
            if (inherited < 550) return 100;
            if (inherited < 750) return 400;
            return 700;
        }
        return TryFontWeight(normalized, out int numericWeight) && numericWeight >= 1 && numericWeight <= 1000
            ? numericWeight
            : inherited;
    }

    private static double ResolveRequestedFontStretch(string value, double inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        switch (normalized) {
            case "": return inherited;
            case "normal": return 100D;
            case "ultra-condensed": return 50D;
            case "extra-condensed": return 62.5D;
            case "condensed": return 75D;
            case "semi-condensed": return 87.5D;
            case "semi-expanded": return 112.5D;
            case "expanded": return 125D;
            case "extra-expanded": return 150D;
            case "ultra-expanded": return 200D;
            case "wider": return NextFontStretch(inherited, wider: true);
            case "narrower": return NextFontStretch(inherited, wider: false);
        }
        if (normalized.EndsWith("%", StringComparison.Ordinal)
            && double.TryParse(
                normalized.Substring(0, normalized.Length - 1),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out double percentage)
            && percentage >= 50D
            && percentage <= 200D) {
            return percentage;
        }
        return inherited;
    }

    private static double NextFontStretch(double inherited, bool wider) {
        double[] values = { 50D, 62.5D, 75D, 87.5D, 100D, 112.5D, 125D, 150D, 200D };
        if (wider) {
            foreach (double value in values) if (value > inherited + 0.0001D) return value;
            return 200D;
        }
        for (int index = values.Length - 1; index >= 0; index--) {
            if (values[index] < inherited - 0.0001D) return values[index];
        }
        return 50D;
    }

    private static void ResolveRequestedFontSlant(
        string tag,
        string value,
        OfficeFontFaceDescriptor inherited,
        out OfficeFontSlant slant,
        out double obliqueAngle) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0) {
            slant = tag == "i" || tag == "em" ? OfficeFontSlant.Italic : inherited.Slant;
            obliqueAngle = slant == OfficeFontSlant.Oblique ? inherited.ObliqueAngleDegrees : 14D;
            return;
        }
        if (normalized == "normal") {
            slant = OfficeFontSlant.Normal;
            obliqueAngle = 14D;
            return;
        }
        if (normalized == "italic") {
            slant = OfficeFontSlant.Italic;
            obliqueAngle = 14D;
            return;
        }
        if (normalized.StartsWith("oblique", StringComparison.Ordinal)) {
            slant = OfficeFontSlant.Oblique;
            obliqueAngle = 14D;
            string angle = normalized.Substring("oblique".Length).Trim();
            if (angle.EndsWith("deg", StringComparison.Ordinal)
                && double.TryParse(angle.Substring(0, angle.Length - 3), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                && parsed > -90D
                && parsed < 90D) {
                obliqueAngle = parsed;
            }
            return;
        }
        slant = inherited.Slant;
        obliqueAngle = inherited.ObliqueAngleDegrees;
    }
}
