using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static bool TryGetRtfHighlightColor(HighlightColorValues value, out byte red, out byte green, out byte blue) {
        red = 0;
        green = 0;
        blue = 0;

        if (value == HighlightColorValues.Black) return true;
        if (value == HighlightColorValues.Blue) {
            blue = 255;
            return true;
        }

        if (value == HighlightColorValues.Cyan) {
            green = 255;
            blue = 255;
            return true;
        }

        if (value == HighlightColorValues.Green) {
            green = 255;
            return true;
        }

        if (value == HighlightColorValues.Magenta) {
            red = 255;
            blue = 255;
            return true;
        }

        if (value == HighlightColorValues.Red) {
            red = 255;
            return true;
        }

        if (value == HighlightColorValues.Yellow) {
            red = 255;
            green = 255;
            return true;
        }

        if (value == HighlightColorValues.DarkBlue) {
            blue = 128;
            return true;
        }

        if (value == HighlightColorValues.DarkCyan) {
            green = 128;
            blue = 128;
            return true;
        }

        if (value == HighlightColorValues.DarkGreen) {
            green = 128;
            return true;
        }

        if (value == HighlightColorValues.DarkMagenta) {
            red = 128;
            blue = 128;
            return true;
        }

        if (value == HighlightColorValues.DarkRed) {
            red = 128;
            return true;
        }

        if (value == HighlightColorValues.DarkYellow) {
            red = 128;
            green = 128;
            return true;
        }

        if (value == HighlightColorValues.DarkGray) {
            red = 128;
            green = 128;
            blue = 128;
            return true;
        }

        if (value == HighlightColorValues.LightGray) {
            red = 192;
            green = 192;
            blue = 192;
            return true;
        }

        return false;
    }

    private static bool TryGetWordHighlightColor(RtfDocument? document, int colorIndex, out HighlightColorValues highlight) {
        highlight = default;
        if (document == null) {
            return false;
        }

        string? colorHex = GetColorHex(document, colorIndex);
        if (colorHex == null) {
            return false;
        }

        switch (colorHex.ToUpperInvariant()) {
            case "000000":
                highlight = HighlightColorValues.Black;
                return true;
            case "0000FF":
                highlight = HighlightColorValues.Blue;
                return true;
            case "00FFFF":
                highlight = HighlightColorValues.Cyan;
                return true;
            case "00FF00":
                highlight = HighlightColorValues.Green;
                return true;
            case "FF00FF":
                highlight = HighlightColorValues.Magenta;
                return true;
            case "FF0000":
                highlight = HighlightColorValues.Red;
                return true;
            case "FFFF00":
                highlight = HighlightColorValues.Yellow;
                return true;
            case "000080":
                highlight = HighlightColorValues.DarkBlue;
                return true;
            case "008080":
                highlight = HighlightColorValues.DarkCyan;
                return true;
            case "008000":
                highlight = HighlightColorValues.DarkGreen;
                return true;
            case "800080":
                highlight = HighlightColorValues.DarkMagenta;
                return true;
            case "800000":
                highlight = HighlightColorValues.DarkRed;
                return true;
            case "808000":
                highlight = HighlightColorValues.DarkYellow;
                return true;
            case "808080":
                highlight = HighlightColorValues.DarkGray;
                return true;
            case "C0C0C0":
                highlight = HighlightColorValues.LightGray;
                return true;
            default:
                return false;
        }
    }
}
