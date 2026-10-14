namespace OfficeIMO.IWork.Internal;

/// <summary>Bounds the shared non-tiered decimal profile for native reading, creation and layout.</summary>
internal static class IWorkDecimalListProfile {
    internal const int MaximumLevel = 8;

    internal static bool IsQualifiedNumberType(ulong value) => value <= 2;

    internal static bool TryGetMarkerStyle(string? sample, out IWorkDecimalListMarkerStyle style) {
        switch (sample) {
            case "1.": style = IWorkDecimalListMarkerStyle.Period; return true;
            case "(1)": style = IWorkDecimalListMarkerStyle.Parentheses; return true;
            case "1)": style = IWorkDecimalListMarkerStyle.ClosingParenthesis; return true;
            default: style = default; return false;
        }
    }

    internal static string FormatMarker(int number, IWorkDecimalListMarkerStyle style) {
        string text = number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return style switch {
            IWorkDecimalListMarkerStyle.Parentheses => "(" + text + ")",
            IWorkDecimalListMarkerStyle.ClosingParenthesis => text + ")",
            _ => text + "."
        };
    }
}
