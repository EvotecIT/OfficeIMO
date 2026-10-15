using AngleSharp.Css.Parser;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string BackgroundDeclarationSentinelPrefix = "-officeimo-internal-background-declaration-";
    private static readonly string[] BackgroundLonghands = {
        "background-image", "background-position", "background-size", "background-repeat",
        "background-attachment", "background-origin", "background-clip", "background-color"
    };

    private static string BackgroundInitialValue(string propertyName) => propertyName.ToLowerInvariant() switch {
        "background-image" => "none",
        "background-position" => "0% 0%",
        "background-size" => "auto",
        "background-repeat" => "repeat",
        "background-attachment" => "scroll",
        "background-origin" => "padding-box",
        "background-clip" => "border-box",
        "background-color" => "transparent",
        _ => string.Empty
    };

    /// <summary>
    /// Expands each layer through the retained grammar, keeping omitted components
    /// in their layer rather than losing them through a synthesized provider list.
    /// The caller applies every component at the authored declaration's priority.
    /// </summary>
    private static bool TryExpandBackgroundShorthand(string value, out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        if (IsCssWideKeyword(value.Trim())) {
            longhands = BackgroundLonghands.Select(name => new KeyValuePair<string, string>(name, value)).ToArray();
            return true;
        }
        IReadOnlyList<string> layers = HtmlRenderCssValues.SplitTopLevelCommas(value);
        var components = BackgroundLonghands.ToDictionary(name => name, _ => new List<string>(), StringComparer.Ordinal);
        for (int index = 0; index < layers.Count; index++) {
            // A CSS-wide keyword applies to the entire shorthand, never one layer.
            if (IsCssWideKeyword(layers[index])) {
                longhands = Array.Empty<KeyValuePair<string, string>>();
                return false;
            }
            string layer = PreserveManagedGradientFunctions(layers[index], preserveBackgroundImages: true);
            layer = PreserveBackgroundProviderLengths(layer, out IReadOnlyDictionary<string, string> originalLengths);
            var parsed = new CssParser().ParseDeclaration("background:" + layer);
            if (parsed == null || parsed.Length == 0) {
                // Drawing owns supported Color 4 syntax. Normalize only a colour
                // component before asking the retained background grammar again.
                IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(layer);
                string normalized = string.Join(" ", tokens.Select(token =>
                    token.IndexOf('(') >= 0 && OfficeIMO.Drawing.OfficeColor.TryParseCss(token, out OfficeIMO.Drawing.OfficeColor color)
                        ? FormatComputedColor(color, color.A / 255D) : token));
                if (normalized != layer) parsed = new CssParser().ParseDeclaration("background:" + normalized);
            }
            if (parsed == null || parsed.Length == 0) {
                longhands = Array.Empty<KeyValuePair<string, string>>();
                return false;
            }
            foreach (string name in BackgroundLonghands) {
                string component = parsed.GetPropertyValue(name);
                if (name == "background-color" && index < layers.Count - 1) {
                    // Only the last layer may contain an authored colour.
                    if (!string.Equals(component, "initial", StringComparison.OrdinalIgnoreCase)) {
                        longhands = Array.Empty<KeyValuePair<string, string>>();
                        return false;
                    }
                    continue;
                }
                if (name == "background-clip" && string.Equals(component, "initial", StringComparison.OrdinalIgnoreCase)) {
                    string origin = parsed.GetPropertyValue("background-origin");
                    // A single visual-box sets both origin and clip.
                    if (!string.Equals(origin, "initial", StringComparison.OrdinalIgnoreCase)) component = origin;
                }
                if (string.IsNullOrWhiteSpace(component) || string.Equals(component, "initial", StringComparison.OrdinalIgnoreCase)) {
                    component = BackgroundInitialValue(name);
                }
                components[name].Add(RestoreBackgroundProviderLengths(RestoreManagedGradientFunctions(component), originalLengths));
            }
        }
        longhands = layers.Count == 0 ? Array.Empty<KeyValuePair<string, string>>()
            : BackgroundLonghands.Select(name => new KeyValuePair<string, string>(name, string.Join(", ", components[name]))).ToArray();
        return layers.Count != 0 && longhands.All(item => IsSupportedBackgroundLonghandSyntax(item.Key, item.Value));
    }

    /// <summary>
    /// Keeps winning initial/unset values explicit so consumers cannot revive an
    /// earlier retained shorthand when a later longhand resets one component.
    /// </summary>
    private static void MaterializeBackgroundResets(IDictionary<string, string> properties, ISet<string> reset) {
        foreach (string name in BackgroundLonghands) {
            if (reset.Contains(name)) properties[name] = BackgroundInitialValue(name);
        }
    }
}
