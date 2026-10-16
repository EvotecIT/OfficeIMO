using AngleSharp.Css.Parser;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static bool IsSupportedBackgroundLonghandSyntax(string name, string value) {
        if (name == "background-color") {
            return value.Equals("currentcolor", StringComparison.OrdinalIgnoreCase)
                || OfficeColor.TryParseCss(value, out _)
                || HtmlCssPropertyParser.Parse("color", value, UnboundedPropertyTokenization).IsAccepted
                || IsRelativeBackgroundColorDiagnosticSyntax(value);
        }
        IReadOnlyList<string> layers = HtmlRenderCssValues.SplitTopLevelCommas(value);
        return layers.Count > 0 && layers.All(layer => {
            IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(layer);
            switch (name) {
                case "background-image": return IsSupportedBackgroundImageSyntax(layer);
                case "background-position": return HtmlCssReplacedElementParser.IsSupportedObjectPositionSyntax(layer);
                case "background-size":
                    return parts.Count == 1 && IsKnownKeyword(parts[0], "cover", "contain")
                        || parts.Count is >= 1 and <= 2 && parts.All(part => part.Equals("auto", StringComparison.OrdinalIgnoreCase)
                            || !part.Equals("normal", StringComparison.OrdinalIgnoreCase) && IsGapComponentSyntax(part));
                case "background-repeat":
                    return parts.Count == 1 && IsKnownKeyword(parts[0], "repeat-x", "repeat-y")
                        || parts.Count is >= 1 and <= 2 && parts.All(part => IsKnownKeyword(part, "repeat", "no-repeat", "round", "space"));
                case "background-attachment": return parts.Count == 1 && IsKnownKeyword(parts[0], "scroll", "fixed", "local");
                case "background-origin":
                case "background-clip": return parts.Count == 1 && IsKnownKeyword(parts[0], "border-box", "padding-box", "content-box");
                default: return false;
            }
        });
    }

    /// <summary>Checks image grammar before opaque provider adaptation, independently of render stop budgets.</summary>
    private static bool IsSupportedBackgroundImageSyntax(string value, int depth = 0) {
        if (depth > 64) return false;
        string image = value.Trim();
        if (image.Equals("none", StringComparison.OrdinalIgnoreCase)) return true;
        int open = image.IndexOf('(');
        if (open <= 0 || !TryFindFunctionEnd(image, open, out int close) || close != image.Length - 1) return false;
        string name = image.Substring(0, open).ToLowerInvariant();
        switch (name) {
            case "linear-gradient":
            case "repeating-linear-gradient": return HtmlCssLinearGradientParser.TryParse(image, int.MaxValue, out _, out _) || ProviderAcceptsBackgroundImage(image);
            case "radial-gradient":
            case "repeating-radial-gradient": return HtmlCssRadialGradientParser.TryParse(image, int.MaxValue, out _, out _) || ProviderAcceptsValidRadialGradient(image);
            case "conic-gradient":
            case "repeating-conic-gradient": return HtmlCssConicGradientParser.TryParse(image, int.MaxValue, out _, out _) || ProviderAcceptsBackgroundImage(image);
            case "image-set":
            case "-webkit-image-set": return HtmlResourcePipeline.IsValidImageSetSyntax(image, child => IsSupportedBackgroundImageSyntax(child, depth + 1));
            case "paint":
                string arguments = image.Substring(open + 1, close - open - 1).Trim();
                int index = 0;
                if (!HtmlCssIdentifierParser.TryRead(arguments, ref index, out string identifier)
                    || IsCssWideKeyword(identifier) || identifier.Equals("default", StringComparison.OrdinalIgnoreCase)) return false;
                string remaining = arguments.Substring(index).Trim();
                return remaining.Length == 0 || remaining[0] == ','
                    && HtmlRenderCssValues.SplitTopLevelCommas(remaining.Substring(1)).All(argument => argument.Length > 0);
            default: return ProviderAcceptsBackgroundImage(image);
        }
    }

    private static bool ProviderAcceptsBackgroundImage(string image) {
        var parsed = new CssParser().ParseDeclaration("background-image:" + image);
        return parsed != null && parsed.Length > 0;
    }

    // Provider acceptance alone does not reject negative radial radii. Validate
    // its descriptor through the existing owned parser while allowing retained
    // colours/interpolation that are admitted only for rendering diagnostics.
    private static bool ProviderAcceptsValidRadialGradient(string image) {
        if (!ProviderAcceptsBackgroundImage(image)) return false;
        int open = image.IndexOf('(');
        IReadOnlyList<string> arguments = HtmlRenderCssValues.SplitTopLevelCommas(image.Substring(open + 1, image.Length - open - 2));
        if (arguments.Count < 2) return false;
        List<string> descriptor = HtmlRenderCssValues.SplitWhitespace(arguments[0]).ToList();
        if (descriptor.Count == 0) return false;
        if (IsSupportedBackgroundLonghandSyntax("background-color", descriptor[0])) return true;
        int interpolation = descriptor.FindIndex(part => part.Equals("in", StringComparison.OrdinalIgnoreCase));
        if (interpolation >= 0) {
            if (interpolation + 1 >= descriptor.Count) return false;
            descriptor.RemoveRange(interpolation, 2);
            if (interpolation + 1 < descriptor.Count
                && IsKnownKeyword(descriptor[interpolation], "shorter", "longer", "increasing", "decreasing")
                && descriptor[interpolation + 1].Equals("hue", StringComparison.OrdinalIgnoreCase)) descriptor.RemoveRange(interpolation, 2);
        }
        string geometry = descriptor.Count == 0 ? "" : string.Join(" ", descriptor) + ",";
        return HtmlCssRadialGradientParser.TryParse("radial-gradient(" + geometry + "red,blue)", int.MaxValue, out _, out _);
    }

    // Relative colours remain diagnostic-only. Admit the known component form
    // without inventing a colour evaluator or accepting arbitrary identifiers.
    private static bool IsRelativeBackgroundColorDiagnosticSyntax(string value) {
        int open = value.IndexOf('(');
        if (open <= 0 || !TryFindFunctionEnd(value, open, out int close) || close != value.Length - 1) return false;
        string function = value.Substring(0, open).Trim().ToLowerInvariant();
        if (!IsKnownKeyword(function, "color", "rgb", "rgba", "hsl", "hsla", "hwb", "lab", "lch", "oklab", "oklch")) return false;
        IReadOnlyList<string> alphaParts = HtmlRenderCssValues.SplitTopLevel(value.Substring(open + 1, close - open - 1), '/');
        if (alphaParts.Count is < 1 or > 2) return false;
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(alphaParts[0]);
        int first = function == "color" ? 3 : 2;
        if (parts.Count != first + 3 || !parts[0].Equals("from", StringComparison.OrdinalIgnoreCase)
            || !OfficeColor.TryParseCss(parts[1], out _) && !HtmlCssPropertyParser.Parse("color", parts[1], UnboundedPropertyTokenization).IsAccepted) return false;
        if (function == "color" && !IsKnownKeyword(parts[2], "srgb", "srgb-linear", "display-p3", "a98-rgb", "prophoto-rgb", "rec2020", "xyz", "xyz-d50", "xyz-d65")) return false;
        string channels = function is "hsl" or "hsla" ? "hsl" : function == "hwb" ? "hwb"
            : function is "lab" or "oklab" ? "lab" : function is "lch" or "oklch" ? "lch"
            : function == "color" && parts[2].StartsWith("xyz", StringComparison.OrdinalIgnoreCase) ? "xyz" : "rgb";
        bool Component(string component) => component.Equals("none", StringComparison.OrdinalIgnoreCase)
            || component.Equals("alpha", StringComparison.OrdinalIgnoreCase)
            || component.Length == 1 && channels.IndexOf(char.ToLowerInvariant(component[0])) >= 0
            || double.TryParse(component.EndsWith("%", StringComparison.Ordinal) ? component.Substring(0, component.Length - 1) : component, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double number) && !double.IsNaN(number) && !double.IsInfinity(number);
        return parts.Skip(first).All(Component)
            && (alphaParts.Count == 1 || HtmlRenderCssValues.SplitWhitespace(alphaParts[1]).Count == 1 && Component(alphaParts[1].Trim()));
    }
}
