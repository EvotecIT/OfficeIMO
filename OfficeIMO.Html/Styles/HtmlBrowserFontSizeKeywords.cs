namespace OfficeIMO.Html;

/// <summary>Browser keyword sizes qualified independently of the Document profile's numeric scales.</summary>
internal static class HtmlBrowserFontSizeKeywords {
    // Observed CSS keyword sizes for integer medium preferences from 9 to 16,
    // ordered xx-small through xxx-large.
    private static readonly double[,] Standards = {
        { 9D, 9D, 9D, 9D, 11D, 14D, 18D, 27D }, // 9px medium
        { 9D, 9D, 9D, 10D, 12D, 15D, 20D, 30D }, // 10px medium
        { 9D, 9D, 10D, 11D, 13D, 17D, 22D, 33D }, // 11px medium
        { 9D, 9D, 10D, 12D, 14D, 18D, 24D, 36D }, // 12px medium
        { 9D, 10D, 12D, 13D, 16D, 20D, 26D, 39D }, // 13px medium
        { 9D, 10D, 12D, 14D, 17D, 21D, 28D, 42D }, // 14px medium
        { 9D, 10D, 13D, 15D, 18D, 23D, 30D, 45D }, // 15px medium
        { 9D, 10D, 13D, 16D, 18D, 24D, 32D, 48D }, // 16px medium
    };
    private static readonly double[,] Quirks = {
        { 9D, 9D, 9D, 9D, 11D, 14D, 18D, 28D }, // 9px medium
        { 9D, 9D, 9D, 10D, 12D, 15D, 20D, 31D }, // 10px medium
        { 9D, 9D, 9D, 11D, 13D, 17D, 22D, 34D }, // 11px medium
        { 9D, 9D, 10D, 12D, 14D, 18D, 24D, 37D }, // 12px medium
        { 9D, 9D, 10D, 13D, 16D, 20D, 26D, 40D }, // 13px medium
        { 9D, 9D, 11D, 14D, 17D, 21D, 28D, 42D }, // 14px medium
        { 9D, 10D, 12D, 15D, 17D, 23D, 30D, 45D }, // 15px medium
        { 9D, 10D, 13D, 16D, 18D, 24D, 32D, 48D }, // 16px medium
    };

    internal static bool TryResolve(string value, double inherited, double medium, bool quirksMode, out double result) {
        int index = value switch {
            "xx-small" => 0,
            "x-small" => 1,
            "small" => 2,
            "medium" => 3,
            "large" => 4,
            "x-large" => 5,
            "xx-large" => 6,
            "xxx-large" => 7,
            _ => -1
        };
        if (index >= 0 && medium >= 9D && medium <= 16D && medium == Math.Truncate(medium)) {
            int row = (int)medium - 9;
            result = (quirksMode ? Quirks : Standards)[row, index];
            return true;
        }
        // Relative keywords use inherited size. Other preferences retain the
        // existing numeric owner; the 20px control independently qualifies that fallback.
        return HtmlRenderCssValues.TryResolveFontSizeKeyword(value, inherited, medium, out result);
    }
}
