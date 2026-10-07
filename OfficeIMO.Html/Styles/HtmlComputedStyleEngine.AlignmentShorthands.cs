namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static readonly string[] PlaceItemsLonghands = { "align-items", "justify-items" };
    private static readonly string[] PlaceSelfLonghands = { "align-self", "justify-self" };
    private static readonly string[] PlaceContentLonghands = { "align-content", "justify-content" };

    private static bool TryExpandAlignmentShorthand(string propertyName, string value,
        out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        string[] names = GetDeferredLayoutShorthandLonghands(propertyName)!;
        string trimmed = value.Trim().ToLowerInvariant();
        if (IsCssWideKeyword(trimmed)) {
            longhands = names.Select(name => new KeyValuePair<string, string>(name, trimmed)).ToArray();
            return true;
        }
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(trimmed);
        int index = 0;
        if (!TryReadAlignmentComponent(parts, ref index, names[0], out string first)) {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }
        string second = propertyName == "place-content" && first.EndsWith("baseline", StringComparison.Ordinal) ? "start" : first;
        if (index < parts.Count && !TryReadAlignmentComponent(parts, ref index, names[1], out second)
            || index != parts.Count) {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }
        longhands = new[] { new KeyValuePair<string, string>(names[0], first), new KeyValuePair<string, string>(names[1], second) };
        return true;
    }

    private static bool TryReadAlignmentComponent(IReadOnlyList<string> parts, ref int index, string propertyName, out string value) {
        value = string.Empty;
        if (index >= parts.Count) return false;
        string token = parts[index++];
        if (token is "first" or "last") {
            if (index >= parts.Count || parts[index++] != "baseline") return false;
            value = token + " baseline";
            return propertyName != "justify-content";
        }
        if (token is "safe" or "unsafe") {
            if (index >= parts.Count || !IsAlignmentPosition(parts[index], propertyName)) return false;
            value = token + " " + parts[index++];
            return true;
        }
        bool content = propertyName.EndsWith("content", StringComparison.Ordinal);
        bool valid = token == "normal" || token == "stretch" || IsAlignmentPosition(token, propertyName)
            || token == "auto" && propertyName.EndsWith("self", StringComparison.Ordinal)
            || token == "baseline" && propertyName != "justify-content"
            || content && token is "space-between" or "space-around" or "space-evenly";
        if (valid) value = token;
        return valid;
    }

    private static bool IsAlignmentPosition(string value, string propertyName) =>
        value is "start" or "end" or "center" or "flex-start" or "flex-end"
        || value is "self-start" or "self-end" && !propertyName.EndsWith("content", StringComparison.Ordinal)
        || value is "left" or "right" && propertyName.StartsWith("justify", StringComparison.Ordinal);
}
