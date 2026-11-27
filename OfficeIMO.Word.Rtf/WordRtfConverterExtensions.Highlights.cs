namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static bool TryGetRtfHighlightColor(WordHighlightColor value, out byte red, out byte green, out byte blue) {
        red = 0;
        green = 0;
        blue = 0;

        if (value == WordHighlightColor.Black) return true;
        if (value == WordHighlightColor.Blue) {
            blue = 255;
            return true;
        }

        if (value == WordHighlightColor.Cyan) {
            green = 255;
            blue = 255;
            return true;
        }

        if (value == WordHighlightColor.Green) {
            green = 255;
            return true;
        }

        if (value == WordHighlightColor.Magenta) {
            red = 255;
            blue = 255;
            return true;
        }

        if (value == WordHighlightColor.Red) {
            red = 255;
            return true;
        }

        if (value == WordHighlightColor.Yellow) {
            red = 255;
            green = 255;
            return true;
        }

        if (value == WordHighlightColor.DarkBlue) {
            blue = 128;
            return true;
        }

        if (value == WordHighlightColor.DarkCyan) {
            green = 128;
            blue = 128;
            return true;
        }

        if (value == WordHighlightColor.DarkGreen) {
            green = 128;
            return true;
        }

        if (value == WordHighlightColor.DarkMagenta) {
            red = 128;
            blue = 128;
            return true;
        }

        if (value == WordHighlightColor.DarkRed) {
            red = 128;
            return true;
        }

        if (value == WordHighlightColor.DarkYellow) {
            red = 128;
            green = 128;
            return true;
        }

        if (value == WordHighlightColor.DarkGray) {
            red = 128;
            green = 128;
            blue = 128;
            return true;
        }

        if (value == WordHighlightColor.LightGray) {
            red = 192;
            green = 192;
            blue = 192;
            return true;
        }

        return false;
    }

    private static bool TryGetWordHighlightColor(RtfDocument? document, int colorIndex, out WordHighlightColor highlight) {
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
                highlight = WordHighlightColor.Black;
                return true;
            case "0000FF":
                highlight = WordHighlightColor.Blue;
                return true;
            case "00FFFF":
                highlight = WordHighlightColor.Cyan;
                return true;
            case "00FF00":
                highlight = WordHighlightColor.Green;
                return true;
            case "FF00FF":
                highlight = WordHighlightColor.Magenta;
                return true;
            case "FF0000":
                highlight = WordHighlightColor.Red;
                return true;
            case "FFFF00":
                highlight = WordHighlightColor.Yellow;
                return true;
            case "000080":
                highlight = WordHighlightColor.DarkBlue;
                return true;
            case "008080":
                highlight = WordHighlightColor.DarkCyan;
                return true;
            case "008000":
                highlight = WordHighlightColor.DarkGreen;
                return true;
            case "800080":
                highlight = WordHighlightColor.DarkMagenta;
                return true;
            case "800000":
                highlight = WordHighlightColor.DarkRed;
                return true;
            case "808000":
                highlight = WordHighlightColor.DarkYellow;
                return true;
            case "808080":
                highlight = WordHighlightColor.DarkGray;
                return true;
            case "C0C0C0":
                highlight = WordHighlightColor.LightGray;
                return true;
            default:
                return false;
        }
    }
}
