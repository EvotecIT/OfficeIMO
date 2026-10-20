using AngleSharp.Css.Parser;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string FontDeclarationSentinelPrefix = "-officeimo-internal-font-declaration-";
    private static readonly string[] FontShorthandLonghands = {
        "font-style", "font-variant", "font-weight", "font-stretch", "font-size", "line-height", "font-family",
        "font-feature-settings", "font-kerning", "font-variant-caps", "font-variant-east-asian",
        "font-variant-ligatures", "font-variant-numeric"
    };

    /// <summary>Expands a complete font shorthand through provider grammar while retaining owned container lengths.</summary>
    internal static bool TryExpandFontShorthand(string value, out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        if (IsCssWideKeyword(value.Trim())) {
            longhands = FontShorthandLonghands.Select(name => new KeyValuePair<string, string>(name, value)).ToArray();
            return true;
        }
        // Use the same CSS parser as stylesheet rules, retaining the declaration's
        // original cascade priority and order when applying its longhands.
        string providerValue = PreserveProviderLengths(value, out IReadOnlyDictionary<string, string> originals,
            expression => ContainsContainerLength(expression) && (expression.IsCalculated || expression.Value >= 0D));
        var parsed = new CssParser().ParseDeclaration("font:" + providerValue);
        var expanded = new List<KeyValuePair<string, string>>();
        if (parsed != null) {
            foreach (string name in FontShorthandLonghands) {
                string resolved = parsed.GetPropertyValue(name);
                if (!string.IsNullOrWhiteSpace(resolved)) expanded.Add(new KeyValuePair<string, string>(name,
                    RestoreProviderLengths(resolved, originals)));
            }
        }
        if (!expanded.Any(item => item.Key == "font-size") || !expanded.Any(item => item.Key == "font-family")) {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }
        foreach (string name in FontShorthandLonghands) {
            if (!expanded.Any(item => item.Key == name)) {
                // Reset-only font subproperties participate at the shorthand's
                // cascade position. Initial also prevents inherited settings
                // from surviving this declaration; font-palette is independent.
                bool resetOnly = name == "font-feature-settings" || name == "font-kerning"
                    || name.StartsWith("font-variant-", StringComparison.Ordinal);
                expanded.Add(new KeyValuePair<string, string>(name, resetOnly ? "initial" : "normal"));
            }
        }
        longhands = expanded;
        return true;
    }

    // Retain only the newer units already handled by the owned numeric grammar;
    // the provider continues to validate the complete font shorthand structure.
    private static bool ContainsContainerLength(HtmlCssMathExpression expression) =>
        expression.Unit == HtmlCssLengthUnit.Cqw || expression.Unit == HtmlCssLengthUnit.Cqh
        || expression.Unit == HtmlCssLengthUnit.Cqi || expression.Unit == HtmlCssLengthUnit.Cqb
        || expression.Unit == HtmlCssLengthUnit.Cqmin || expression.Unit == HtmlCssLengthUnit.Cqmax
        || expression.Children.Any(ContainsContainerLength);

    private static void ResolveDeferredFontLonghands(Dictionary<string, string> properties, ISet<string> deferred,
        IReadOnlyDictionary<string, string>? parentProperties, ISet<string> inherited, ISet<string> reset, ISet<string> specified,
        bool enforceResolutionLimits) {
        foreach (string name in deferred) {
            string value = "unset";
            if (HtmlCssCustomPropertyResolver.TryResolve(properties[name],
                    customName => properties.TryGetValue(customName, out string? local) ? local
                        : parentProperties != null && parentProperties.TryGetValue(customName, out string? parent) ? parent : null,
                    out string shorthand, enforceResolutionLimits)
                && TryExpandFontShorthand(shorthand, out IReadOnlyList<KeyValuePair<string, string>> longhands)) {
                value = longhands.First(item => item.Key == name).Value;
            }
            // Invalid substitution still occupies the shorthand's cascade position.
            // Resolve its longhands as unset rather than reviving earlier declarations.
            CssKeywordResolution resolved = ResolveCssWideKeyword(name, value, parentProperties);
            if (resolved.HasValue) {
                properties[name] = resolved.Value;
                reset.Remove(name);
                if (resolved.InheritsComputedValue) {
                    inherited.Add(name);
                    specified.Remove(name);
                } else {
                    inherited.Remove(name);
                    specified.Add(name);
                }
            } else {
                properties.Remove(name);
                inherited.Remove(name);
                specified.Remove(name);
                reset.Add(name);
            }
        }
    }

}
