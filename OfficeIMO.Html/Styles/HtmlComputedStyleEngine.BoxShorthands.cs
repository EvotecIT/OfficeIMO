namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static readonly string[] PhysicalBoxSides = { "top", "right", "bottom", "left" };

    private static readonly string[] PhysicalBoxShorthands = { "margin", "padding", "border", "border-width", "border-style", "border-color" };
    private static readonly string[] MarginLonghands = { "margin-top", "margin-right", "margin-bottom", "margin-left" };
    private static readonly string[] PaddingLonghands = { "padding-top", "padding-right", "padding-bottom", "padding-left" };
    private static readonly string[] BorderWidthLonghands = { "border-top-width", "border-right-width", "border-bottom-width", "border-left-width" };
    private static readonly string[] BorderStyleLonghands = { "border-top-style", "border-right-style", "border-bottom-style", "border-left-style" };
    private static readonly string[] BorderColorLonghands = { "border-top-color", "border-right-color", "border-bottom-color", "border-left-color" };

    private static bool TryExpandPhysicalBoxShorthand(
        string propertyName,
        string value,
        out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        string normalizedName = propertyName.Trim().ToLowerInvariant();
        if (normalizedName == "border") {
            string width, style, color;
            if (IsCssWideKeyword(value.Trim())) {
                width = style = color = value;
            } else if (!TryExpandBorderComponents(value, out width, out style, out color)) {
                longhands = Array.Empty<KeyValuePair<string, string>>();
                return false;
            }
            var border = new KeyValuePair<string, string>[12];
            for (int index = 0; index < 4; index++) {
                border[index * 3] = new KeyValuePair<string, string>(BorderWidthLonghands[index], width);
                border[index * 3 + 1] = new KeyValuePair<string, string>(BorderStyleLonghands[index], style);
                border[index * 3 + 2] = new KeyValuePair<string, string>(BorderColorLonghands[index], color);
            }
            longhands = border;
            return true;
        }

        string[] names;
        switch (normalizedName) {
            case "margin": names = MarginLonghands; break;
            case "padding": names = PaddingLonghands; break;
            case "border-width": names = BorderWidthLonghands; break;
            case "border-style": names = BorderStyleLonghands; break;
            case "border-color": names = BorderColorLonghands; break;
            default:
                longhands = Array.Empty<KeyValuePair<string, string>>();
                return false;
        }

        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4) {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }
        longhands = new[] {
            new KeyValuePair<string, string>(names[0], tokens[0]),
            new KeyValuePair<string, string>(names[1], tokens.Count > 1 ? tokens[1] : tokens[0]),
            new KeyValuePair<string, string>(names[2], tokens.Count > 2 ? tokens[2] : tokens[0]),
            new KeyValuePair<string, string>(names[3], tokens.Count > 3 ? tokens[3] : tokens.Count > 1 ? tokens[1] : tokens[0])
        };
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

        foreach (string shorthand in PhysicalBoxShorthands) {
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
