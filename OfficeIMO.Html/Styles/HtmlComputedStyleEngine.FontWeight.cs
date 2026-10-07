using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static void ResolveRelativeComputedFontWeight(
        IDictionary<string, string> properties,
        IReadOnlyDictionary<string, string>? parentProperties) {
        if (!properties.TryGetValue("font-weight", out string? value)) return;
        value = value.Trim();
        if (!(string.Equals(value, "bolder", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "lighter", StringComparison.OrdinalIgnoreCase))) return;

        int inheritedWeight = 400;
        if (parentProperties != null && parentProperties.TryGetValue("font-weight", out string? parentWeight)
            && OfficeFontFaceCssParser.TryWeight(parentWeight, 400, out int parsedParentWeight)) {
            inheritedWeight = parsedParentWeight;
        }
        // Relative keywords become an absolute computed value before descendants
        // inherit it or the inline SVG adapter projects it into a new reader.
        if (OfficeFontFaceCssParser.TryWeight(value, inheritedWeight, out int weight)) {
            properties["font-weight"] = weight.ToString(CultureInfo.InvariantCulture);
        }
    }
}
