using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Parses reusable pieces of AcroForm default appearance strings used by generated widget appearances.
/// </summary>
internal static class PdfDefaultAppearanceParser {
    private static readonly char[] Separators = { ' ', '\t', '\r', '\n' };

    public static bool TryReadTextColor(string? defaultAppearance, out PdfColor color) {
        color = PdfColor.Black;
        if (string.IsNullOrWhiteSpace(defaultAppearance)) {
            return false;
        }

        string[] tokens = defaultAppearance!.Split(Separators, StringSplitOptions.RemoveEmptyEntries);
        bool found = false;
        for (int i = 0; i < tokens.Length; i++) {
            if (string.Equals(tokens[i], "g", StringComparison.Ordinal) &&
                i >= 1 &&
                TryReadNumber(tokens[i - 1], out double gray)) {
                color = FromGray(gray);
                found = true;
                continue;
            }

            if (string.Equals(tokens[i], "rg", StringComparison.Ordinal) &&
                i >= 3 &&
                TryReadNumber(tokens[i - 3], out double red) &&
                TryReadNumber(tokens[i - 2], out double green) &&
                TryReadNumber(tokens[i - 1], out double blue)) {
                color = new PdfColor(ClampColor(red), ClampColor(green), ClampColor(blue));
                found = true;
                continue;
            }

            if (string.Equals(tokens[i], "k", StringComparison.Ordinal) &&
                i >= 4 &&
                TryReadNumber(tokens[i - 4], out double cyan) &&
                TryReadNumber(tokens[i - 3], out double magenta) &&
                TryReadNumber(tokens[i - 2], out double yellow) &&
                TryReadNumber(tokens[i - 1], out double black)) {
                color = FromCmyk(cyan, magenta, yellow, black);
                found = true;
            }
        }

        return found;
    }

    public static bool TryReadFontSize(string? defaultAppearance, out double fontSize) {
        fontSize = 0D;
        if (string.IsNullOrWhiteSpace(defaultAppearance)) {
            return false;
        }

        string[] tokens = defaultAppearance!.Split(Separators, StringSplitOptions.RemoveEmptyEntries);
        bool found = false;
        for (int i = 0; i < tokens.Length; i++) {
            if (string.Equals(tokens[i], "Tf", StringComparison.Ordinal) &&
                i >= 2 &&
                TryReadNumber(tokens[i - 1], out double parsedFontSize) &&
                parsedFontSize > 0D) {
                fontSize = parsedFontSize;
                found = true;
            }
        }

        return found;
    }

    public static bool TryReadFontResourceName(string? defaultAppearance, out string fontResourceName) {
        fontResourceName = string.Empty;
        if (string.IsNullOrWhiteSpace(defaultAppearance)) {
            return false;
        }

        string[] tokens = defaultAppearance!.Split(Separators, StringSplitOptions.RemoveEmptyEntries);
        bool found = false;
        for (int i = 0; i < tokens.Length; i++) {
            if (string.Equals(tokens[i], "Tf", StringComparison.Ordinal) &&
                i >= 2 &&
                tokens[i - 2].Length > 1 &&
                tokens[i - 2][0] == '/') {
                fontResourceName = tokens[i - 2].Substring(1);
                found = true;
            }
        }

        return found;
    }

    private static bool TryReadNumber(string token, out double value) =>
        double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
        !double.IsNaN(value) &&
        !double.IsInfinity(value);

    private static PdfColor FromGray(double gray) {
        double component = ClampColor(gray);
        return new PdfColor(component, component, component);
    }

    private static PdfColor FromCmyk(double cyan, double magenta, double yellow, double black) {
        double key = ClampColor(black);
        return new PdfColor(
            (1D - ClampColor(cyan)) * (1D - key),
            (1D - ClampColor(magenta)) * (1D - key),
            (1D - ClampColor(yellow)) * (1D - key));
    }

    private static double ClampColor(double value) {
        if (value < 0D) {
            return 0D;
        }

        return value > 1D ? 1D : value;
    }
}
