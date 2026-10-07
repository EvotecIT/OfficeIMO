namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static readonly string[] TextDecorationLonghands = {
        "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-decoration-thickness"
    };

    private static bool TryExpandTextDecorationShorthand(
        string value,
        out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        string line = "none";
        string style = "solid";
        string color = "currentcolor";
        string thickness = "auto";
        string normalized = value.Trim();
        if (IsCssWideKeyword(normalized)) {
            line = style = color = thickness = normalized;
        } else {
            IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(normalized);
            if (tokens.Count == 0 || tokens.Count > 6) {
                longhands = Array.Empty<KeyValuePair<string, string>>();
                return false;
            }

            var lines = new List<string>(3);
            bool styleSet = false;
            bool colorSet = false;
            bool thicknessSet = false;
            foreach (string token in tokens) {
                string lower = token.ToLowerInvariant();
                if (IsKnownKeyword(lower, "none", "underline", "overline", "line-through")) {
                    if (lines.Count == 3 || lines.Contains(lower)
                        || lower == "none" && lines.Count > 0 || lines.Contains("none")) {
                        longhands = Array.Empty<KeyValuePair<string, string>>();
                        return false;
                    }
                    lines.Add(lower);
                } else if (!styleSet && IsKnownKeyword(lower, "solid", "double", "dotted", "dashed", "wavy")) {
                    style = lower;
                    styleSet = true;
                } else if (!colorSet && (lower == "currentcolor" || HtmlRenderCssValues.TryColor(token, out _))) {
                    color = token;
                    colorSet = true;
                } else if (!thicknessSet && IsTextDecorationThicknessSyntax(token)) {
                    thickness = token;
                    thicknessSet = true;
                } else {
                    longhands = Array.Empty<KeyValuePair<string, string>>();
                    return false;
                }
            }
            if (lines.Count > 0) line = string.Join(" ", lines);
        }

        longhands = new[] {
            new KeyValuePair<string, string>("text-decoration-line", line),
            new KeyValuePair<string, string>("text-decoration-style", style),
            new KeyValuePair<string, string>("text-decoration-color", color),
            new KeyValuePair<string, string>("text-decoration-thickness", thickness)
        };
        return true;
    }

    private static bool IsTextDecorationThicknessSyntax(string value) =>
        IsKnownKeyword(value.ToLowerInvariant(), "auto", "from-font")
        || HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: true, allowUnitlessZero: true)
            && TryValidateCssLength(value, out _);
}
