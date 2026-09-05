namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static readonly string[] PhysicalBoxSides = { "top", "right", "bottom", "left" };

    private static bool TryExpandPhysicalBoxShorthand(
        string propertyName,
        string value,
        out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        string normalizedName = propertyName.Trim().ToLowerInvariant();
        if (normalizedName == "border") {
            if (IsCssWideKeyword(value.Trim())) {
                longhands = PhysicalBoxSides
                    .SelectMany(side => new[] {
                        new KeyValuePair<string, string>("border-" + side + "-width", value),
                        new KeyValuePair<string, string>("border-" + side + "-style", value),
                        new KeyValuePair<string, string>("border-" + side + "-color", value)
                    })
                    .ToArray();
                return true;
            }
            if (!TryExpandBorderComponents(value, out string width, out string style, out string color)) {
                longhands = Array.Empty<KeyValuePair<string, string>>();
                return false;
            }
            longhands = PhysicalBoxSides
                .SelectMany(side => new[] {
                    new KeyValuePair<string, string>("border-" + side + "-width", width),
                    new KeyValuePair<string, string>("border-" + side + "-style", style),
                    new KeyValuePair<string, string>("border-" + side + "-color", color)
                })
                .ToArray();
            return true;
        }

        string prefix;
        if (normalizedName == "margin" || normalizedName == "padding") {
            prefix = normalizedName + "-";
        } else if (normalizedName == "border-width") {
            prefix = "border-";
        } else if (normalizedName == "border-style") {
            prefix = "border-";
        } else if (normalizedName == "border-color") {
            prefix = "border-";
        } else {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }

        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4) {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }

        string[] expanded = {
            tokens[0],
            tokens.Count > 1 ? tokens[1] : tokens[0],
            tokens.Count > 2 ? tokens[2] : tokens[0],
            tokens.Count > 3 ? tokens[3] : tokens.Count > 1 ? tokens[1] : tokens[0]
        };
        string suffix = normalizedName == "border-width" ? "-width"
            : normalizedName == "border-style" ? "-style"
            : normalizedName == "border-color" ? "-color"
            : string.Empty;
        longhands = PhysicalBoxSides
            .Select((side, index) => new KeyValuePair<string, string>(prefix + side + suffix, expanded[index]))
            .ToArray();
        return true;
    }

    private static bool TryExpandBorderComponents(string value, out string width, out string style, out string color) {
        width = "medium";
        style = "none";
        color = "currentcolor";
        bool widthSet = false;
        bool styleSet = false;
        bool colorSet = false;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 3) return false;
        foreach (string token in tokens) {
            if (!widthSet && HtmlCssBoxStrokeParser.IsSupportedSideWidthSyntax(token)) {
                width = token;
                widthSet = true;
            } else if (!styleSet && HtmlCssBoxStrokeParser.IsSupportedSideStyleSyntax(token)) {
                style = token;
                styleSet = true;
            } else if (!colorSet && HtmlCssBoxStrokeParser.IsSupportedSideColorSyntax(token)) {
                color = token;
                colorSet = true;
            } else {
                return false;
            }
        }
        return true;
    }

    private static void ExpandResolvedPhysicalBoxShorthands(
        Dictionary<string, string> properties,
        Dictionary<string, HtmlCssCascadePriority> priorities,
        ISet<string> inherited,
        ISet<string> reset,
        ISet<string> specified) {
        string[] shorthands = { "margin", "padding", "border", "border-width", "border-style", "border-color" };
        foreach (string shorthand in shorthands) {
            if (!properties.TryGetValue(shorthand, out string? value)
                || !TryExpandPhysicalBoxShorthand(shorthand, value, out IReadOnlyList<KeyValuePair<string, string>> longhands)) {
                continue;
            }

            foreach (KeyValuePair<string, string> longhand in longhands) {
                if (properties.ContainsKey(longhand.Key)
                    && (!priorities.TryGetValue(shorthand, out HtmlCssCascadePriority candidate)
                        || priorities.TryGetValue(longhand.Key, out HtmlCssCascadePriority existing)
                        && !candidate.OutranksOrEquals(existing))) {
                    continue;
                }

                properties[longhand.Key] = longhand.Value;
                if (priorities.TryGetValue(shorthand, out HtmlCssCascadePriority priority)) priorities[longhand.Key] = priority;
                if (inherited.Contains(shorthand)) inherited.Add(longhand.Key); else inherited.Remove(longhand.Key);
                if (specified.Contains(shorthand)) specified.Add(longhand.Key); else specified.Remove(longhand.Key);
                reset.Remove(longhand.Key);
            }
        }
    }
}
