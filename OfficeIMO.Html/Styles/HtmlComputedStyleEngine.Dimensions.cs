namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static readonly HashSet<string> DimensionDeclarationNames = new(StringComparer.OrdinalIgnoreCase) {
        "width", "height", "min-width", "min-height", "max-width", "max-height"
    };

    private static bool IsIntrinsicDimensionSyntax(string value) {
        string normalized = StripCssCommentsOutsideStrings(value).Trim();
        if (IsKnownKeyword(normalized, "min-content", "max-content", "fit-content")) return true;
        if (!normalized.StartsWith("fit-content(", StringComparison.OrdinalIgnoreCase)
            || !normalized.EndsWith(")", StringComparison.Ordinal)) return false;
        string argument = normalized.Substring(12, normalized.Length - 13);
        OfficeIMO.Html.Css.HtmlCssPropertyParseResult parsed = OfficeIMO.Html.Css.HtmlCssPropertyParser.Parse(
            "min-width", argument, UnboundedPropertyTokenization);
        return parsed.IsAccepted && parsed.Value?.Kind is OfficeIMO.Html.Css.HtmlCssPropertyValueKind.Length
            or OfficeIMO.Html.Css.HtmlCssPropertyValueKind.Percentage or OfficeIMO.Html.Css.HtmlCssPropertyValueKind.Calculation;
    }
}
