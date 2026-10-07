using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    // Authored layer indices are nonnegative. This implicit first layer places
    // normal SVG presentation hints before every authored stylesheet layer.
    private static readonly CascadeLayerOrder SvgPresentationAttributeLayerOrder = new CascadeLayerOrder(new[] { -1 });
    // These inherited SVG text and paint properties must participate in the
    // cascade before computed styles are projected into the shared SVG reader.
    // Geometry and transforms retain their separate SVG parsing contract.
    private static readonly HashSet<string> SvgTextAndPaintPresentationProperties = new HashSet<string>(StringComparer.Ordinal) {
        "color", "direction", "unicode-bidi", "display", "visibility", "font-family", "font-size",
        "font-stretch", "font-style", "font-variant", "font-weight", "letter-spacing",
        "word-spacing", "white-space", "writing-mode", "fill", "fill-opacity", "fill-rule",
        "stroke", "stroke-width", "stroke-opacity", "stroke-dasharray", "stroke-dashoffset",
        "stroke-linecap", "stroke-linejoin", "stroke-miterlimit", "marker-start", "marker-mid",
        "marker-end", "text-anchor", "dominant-baseline", "baseline-shift"
    };

    private static void ApplySvgTextAndPaintPresentationAttributes(
        IElement element,
        IReadOnlyDictionary<string, string>? parentProperties,
        IDictionary<string, CascadedProperty> properties,
        HtmlCssProcessingBudget budget) {
        if (!string.Equals(element.NamespaceUri, "http://www.w3.org/2000/svg", StringComparison.Ordinal)) return;

        foreach (IAttr attribute in element.Attributes) {
            string name = attribute.Name;
            if (!string.IsNullOrEmpty(attribute.NamespaceUri)
                || !SvgTextAndPaintPresentationProperties.Contains(name)) continue;
            // Animation elements use fill to control the animation lifecycle.
            if (name == "fill" && element.LocalName is "animate" or "animateMotion" or "animateTransform" or "set") continue;
            budget.RecordDeclaration();
            string value = StripTrailingImportant(attribute.Value.Trim(), out bool isImportant);
            if (isImportant) continue; // A presentation attribute is a value, not a declaration.
            if (name == "font-size"
                && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double size)
                && !double.IsNaN(size) && !double.IsInfinity(size) && size >= 0D) {
                // SVG unitless font sizes are user-unit lengths, rather than CSS numbers.
                value += "px";
            }
            ApplyDeclaration(properties, parentProperties, name, value, false,
                Specificity.PresentationalHint, -1, layerOrder: SvgPresentationAttributeLayerOrder,
                enforceResolutionLimits: budget.HasDeclarationLimit);
        }
    }
}
