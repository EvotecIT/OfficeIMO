namespace OfficeIMO.Html.Rtf;

internal static partial class HtmlStyleDeclarationParser {
    private static RtfCapsStyle? ParseTextTransform(string value) {
        switch (value) {
            case "uppercase":
                return RtfCapsStyle.Caps;
            case "none":
                return RtfCapsStyle.None;
            default:
                return null;
        }
    }

    private static RtfCapsStyle? ParseFontVariantCaps(string value) {
        string normalized = " " + value.Replace('-', ' ') + " ";
        if (ContainsWord(normalized, "small caps")) {
            return RtfCapsStyle.SmallCaps;
        }

        return value == "normal" ? RtfCapsStyle.None : null;
    }

    private static RtfCapsStyle? ParseRtfCapsStyle(string value) {
        switch (value) {
            case "none":
                return RtfCapsStyle.None;
            case "caps":
                return RtfCapsStyle.Caps;
            case "small-caps":
                return RtfCapsStyle.SmallCaps;
            default:
                return null;
        }
    }
}
