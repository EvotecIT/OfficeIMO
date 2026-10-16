using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    /// <summary>Returns the footnote-call font when CSS does not override its user-agent default.</summary>
    internal static double ResolveDefaultFootnoteCallFontSize(double parentFontSize) => parentFontSize * 0.7D;

    private double ResolveComputedFontSize(
        string tag,
        HtmlComputedStyle computed,
        string fontSizeValue,
        double parentFontSize,
        bool pseudoElement,
        string pseudoSemanticRole) {
        // Call defaults must precede line-height and other font-relative properties.
        // Explicit inherit/reset declarations retain the ordinary CSS font contract.
        if (pseudoElement && pseudoSemanticRole == "footnote-call" && !HasAuthoredValue(computed, "font-size")) {
            return ResolveDefaultFootnoteCallFontSize(parentFontSize);
        }
        double defaultFontSize = pseudoElement ? parentFontSize : ResolveDefaultTagFontSize(tag, parentFontSize);
        return computed.IsOriginRevertedValue("font-size")
            ? defaultFontSize
            : computed.IsResetValue("font-size")
            ? _options.DefaultFontSize
            : computed.IsImplicitlyInheritedValue("font-size")
            ? defaultFontSize
            : computed.IsInheritedValue("font-size")
            ? parentFontSize
            : string.IsNullOrWhiteSpace(fontSizeValue)
            ? defaultFontSize
            : ResolveFontSize(fontSizeValue, parentFontSize,
                tag == "html" && !pseudoElement ? _options.DefaultFontSize : _rootFontSize);
    }

    internal static OfficeFontFaceDescriptor ResolveFontFaceDescriptor(
        string tag,
        HtmlComputedStyle computed,
        OfficeFontFaceDescriptor inherited) {
        // HTML defaults are materialized by the computed-style owner before
        // inheritance, including the normal values for explicit initial resets.
        string weightValue = computed.IsImplicitlyInheritedValue("font-weight")
            ? string.Empty
            : computed.GetValue("font-weight");
        if (tag == "math" && (!HasAuthoredValue(computed, "font-weight") || computed.IsResetValue("font-weight")))
            weightValue = "normal";
        int weight = OfficeFontFaceCssParser.TryWeight(weightValue, inherited.Weight, out int parsedWeight) ? parsedWeight : inherited.Weight;
        double stretch = OfficeFontFaceCssParser.TryStretch(computed.GetValue("font-stretch"), inherited.StretchPercent, out double parsedStretch)
            ? parsedStretch : inherited.StretchPercent;
        string slantValue = tag == "math" && (!HasAuthoredValue(computed, "font-style") || computed.IsResetValue("font-style"))
            ? "normal" : computed.GetValue("font-style");
        if (!OfficeFontFaceCssParser.TrySlant(slantValue, inherited, out OfficeFontSlant slant, out double angle)) {
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
