using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    internal static OfficeFontFaceDescriptor ResolveFontFaceDescriptor(
        string tag,
        HtmlComputedStyle computed,
        OfficeFontFaceDescriptor inherited) {
        // User-agent bold defaults apply after implicit inheritance, but an authored
        // font-weight on the element still overrides the default.
        string weightValue = computed.IsImplicitlyInheritedValue("font-weight")
            ? string.Empty
            : computed.GetValue("font-weight");
        if (tag == "math" && (!HasAuthoredValue(computed, "font-weight") || computed.IsResetValue("font-weight")))
            weightValue = "normal";
        bool heading = tag.Length == 2 && tag[0] == 'h' && tag[1] >= '1' && tag[1] <= '6';
        int defaultWeight = weightValue.Trim().Length == 0 && (heading || tag == "th" || tag == "b" || tag == "strong") ? 700 : inherited.Weight;
        int weight = OfficeFontFaceCssParser.TryWeight(weightValue, defaultWeight, out int parsedWeight) ? parsedWeight : inherited.Weight;
        double stretch = OfficeFontFaceCssParser.TryStretch(computed.GetValue("font-stretch"), inherited.StretchPercent, out double parsedStretch)
            ? parsedStretch : inherited.StretchPercent;
        string slantValue = tag == "math" && (!HasAuthoredValue(computed, "font-style") || computed.IsResetValue("font-style"))
            ? "normal" : computed.GetValue("font-style");
        OfficeFontFaceDescriptor slantDefault = slantValue.Trim().Length == 0 && (tag == "i" || tag == "em")
            ? new OfficeFontFaceDescriptor(inherited.Weight, inherited.StretchPercent, OfficeFontSlant.Italic) : inherited;
        if (!OfficeFontFaceCssParser.TrySlant(slantValue, slantDefault, out OfficeFontSlant slant, out double angle)) {
            slant = inherited.Slant;
            angle = inherited.ObliqueAngleDegrees;
        }
        return new OfficeFontFaceDescriptor(weight, stretch, slant, angle);
    }

    internal static string ResolveFontFamily(string tag, HtmlComputedStyle computed, string inherited) {
        if (tag == "math") {
            if (!HasAuthoredValue(computed, "font-family")) return "math";
            if (computed.IsResetValue("font-family")) return "serif";
            if (computed.IsInheritedValue("font-family")) return inherited;
        }
        return HtmlRenderCssValues.FontFamilyList(computed.GetValue("font-family"), inherited);
    }

}
