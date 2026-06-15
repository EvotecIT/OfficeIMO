namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlStyleDeclarationParser {
    private static int? ParseCharacterSpacing(string value) {
        if (value == "normal") {
            return 0;
        }

        return ParseTwips(value);
    }

    private static int? ParseCharacterScale(string value) {
        switch (value) {
            case "normal":
            case "initial":
                return 100;
            case "ultra-condensed":
                return 50;
            case "extra-condensed":
                return 62;
            case "condensed":
                return 75;
            case "semi-condensed":
                return 87;
            case "semi-expanded":
                return 112;
            case "expanded":
                return 125;
            case "extra-expanded":
                return 150;
            case "ultra-expanded":
                return 200;
        }

        if (!value.EndsWith("%", StringComparison.Ordinal)) {
            return null;
        }

        return int.TryParse(value.Substring(0, value.Length - 1), out int percent) && percent > 0 ? percent : null;
    }

    private static int? ParseCharacterOffset(string value) {
        switch (value) {
            case "baseline":
            case "middle":
                return 0;
            case "super":
            case "text-top":
            case "sub":
            case "text-bottom":
                return null;
        }

        int? twips = ParseTwips(value);
        return twips.HasValue ? (int)Math.Round(twips.Value / 10d, MidpointRounding.AwayFromZero) : null;
    }

    private static int? ParseRtfCharacterScale(string value) =>
        int.TryParse(value, out int percent) && percent > 0 ? percent : null;

    private static int? ParseRtfCharacterOffset(string value) =>
        int.TryParse(value, out int halfPoints) ? halfPoints : null;
}
