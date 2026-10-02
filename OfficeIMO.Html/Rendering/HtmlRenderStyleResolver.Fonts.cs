using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private static OfficeFontFaceDescriptor ResolveFontFaceDescriptor(
        string tag, HtmlComputedStyle computed, OfficeFontFaceDescriptor inherited) {
        string weightValue = computed.GetValue("font-weight");
        bool heading = tag.Length == 2 && tag[0] == 'h' && tag[1] >= '1' && tag[1] <= '6';
        int defaultWeight = weightValue.Trim().Length == 0 && (heading || tag == "b" || tag == "strong") ? 700 : inherited.Weight;
        int weight = OfficeFontFaceCssParser.TryWeight(weightValue, defaultWeight, out int parsedWeight) ? parsedWeight : inherited.Weight;
        double stretch = OfficeFontFaceCssParser.TryStretch(computed.GetValue("font-stretch"), inherited.StretchPercent, out double parsedStretch)
            ? parsedStretch : inherited.StretchPercent;
        string slantValue = computed.GetValue("font-style");
        OfficeFontFaceDescriptor slantDefault = slantValue.Trim().Length == 0 && (tag == "i" || tag == "em")
            ? new OfficeFontFaceDescriptor(inherited.Weight, inherited.StretchPercent, OfficeFontSlant.Italic) : inherited;
        if (!OfficeFontFaceCssParser.TrySlant(slantValue, slantDefault, out OfficeFontSlant slant, out double angle)) {
            slant = inherited.Slant;
            angle = inherited.ObliqueAngleDegrees;
        }
        return new OfficeFontFaceDescriptor(weight, stretch, slant, angle);
    }
}
