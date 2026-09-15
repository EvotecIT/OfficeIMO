using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static void NormalizeOwnedComputedValues(IDictionary<string, string> properties) {
        if (properties.TryGetValue("opacity", out string? opacityText)) {
            HtmlCssPropertyParseResult opacity = HtmlCssPropertyParser.Parse("opacity", opacityText, UnboundedPropertyTokenization);
            HtmlCssNumericValue? numeric = opacity.Value?.NumericValue;
            if (opacity.Status == HtmlCssPropertyParseStatus.Parsed && numeric != null) {
                double value = numeric.Type == HtmlCssNumericType.Percentage ? numeric.Value / 100D : numeric.Value;
                value = value < 0D ? 0D : value > 1D ? 1D : value;
                properties["opacity"] = value.ToString("R", CultureInfo.InvariantCulture);
            }
        }

        if (properties.TryGetValue("color", out string? colorText)) {
            HtmlCssPropertyParseResult parsed = HtmlCssPropertyParser.Parse("color", colorText, UnboundedPropertyTokenization);
            HtmlCssColorFunctionValue? function = parsed.Value?.ColorFunction;
            if (parsed.Value?.Kind == HtmlCssPropertyValueKind.ColorFunction
                && function != null
                && TryResolveOwnedColor(function, out OfficeColor color)) {
                properties["color"] = FormatComputedColor(color, ResolveAlpha(function.Alpha));
            }
        }
    }

    private static bool TryResolveOwnedColor(HtmlCssColorFunctionValue function, out OfficeColor color) {
        string name = function.Kind == HtmlCssColorFunctionKind.Rgb ? "rgb"
            : function.Kind == HtmlCssColorFunctionKind.Hsl ? "hsl"
            : "hwb";
        string[] components = new string[function.Components.Count];
        for (int i = 0; i < components.Length; i++) {
            HtmlCssColorComponent component = function.Components[i];
            if (component.Kind == HtmlCssColorComponentKind.None) {
                components[i] = i == 0 ? "0" : function.Kind == HtmlCssColorFunctionKind.Rgb ? "0" : "0%";
                continue;
            }

            string number = component.Value!.Value.ToString("R", CultureInfo.InvariantCulture);
            if (component.Kind == HtmlCssColorComponentKind.Percentage
                || (function.Kind != HtmlCssColorFunctionKind.Rgb && i > 0)) {
                number += "%";
            } else if (component.Kind == HtmlCssColorComponentKind.Angle) {
                number += "deg";
            }
            components[i] = number;
        }

        return OfficeColor.TryParseCss(name + "(" + string.Join(" ", components) + ")", out color);
    }

    private static string FormatComputedColor(OfficeColor color, double alpha) =>
        "rgba(" + color.R.ToString(CultureInfo.InvariantCulture)
        + ", " + color.G.ToString(CultureInfo.InvariantCulture)
        + ", " + color.B.ToString(CultureInfo.InvariantCulture)
        + ", " + alpha.ToString("R", CultureInfo.InvariantCulture) + ")";

    private static double ResolveAlpha(HtmlCssColorComponent alpha) {
        if (alpha.Kind == HtmlCssColorComponentKind.None) return 0D;
        double value = alpha.Value ?? 1D;
        if (alpha.Kind == HtmlCssColorComponentKind.Percentage) value /= 100D;
        return value < 0D ? 0D : value > 1D ? 1D : value;
    }
}
