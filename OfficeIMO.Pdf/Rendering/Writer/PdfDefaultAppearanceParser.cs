using System.Globalization;
using System.Threading;

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

    internal static bool TryReadTextColor(string? defaultAppearance, out PdfColor color, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled) return TryReadTextColor(defaultAppearance, out color);
        color = PdfColor.Black;
        if (Guard.IsNullOrWhiteSpaceCancellable(defaultAppearance, cancellationToken)) return false;

        bool found = false;
        string? previous1 = null, previous2 = null, previous3 = null, previous4 = null;
        foreach (string? token in EnumerateTokensCancellable(defaultAppearance!, cancellationToken)) {
            if (token == "g" && previous1 is not null && TryReadNumber(previous1, out double gray)) {
                color = FromGray(gray);
                found = true;
            } else if (token == "rg" && previous3 is not null && previous2 is not null && previous1 is not null &&
                       TryReadNumber(previous3, out double red) && TryReadNumber(previous2, out double green) &&
                       TryReadNumber(previous1, out double blue)) {
                color = new PdfColor(ClampColor(red), ClampColor(green), ClampColor(blue));
                found = true;
            } else if (token == "k" && previous4 is not null && previous3 is not null && previous2 is not null && previous1 is not null &&
                       TryReadNumber(previous4, out double cyan) && TryReadNumber(previous3, out double magenta) &&
                       TryReadNumber(previous2, out double yellow) && TryReadNumber(previous1, out double black)) {
                color = FromCmyk(cyan, magenta, yellow, black);
                found = true;
            }
            previous4 = previous3;
            previous3 = previous2;
            previous2 = previous1;
            previous1 = token;
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

    internal static bool TryReadFontSize(string? defaultAppearance, out double fontSize, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled) return TryReadFontSize(defaultAppearance, out fontSize);
        fontSize = 0D;
        if (Guard.IsNullOrWhiteSpaceCancellable(defaultAppearance, cancellationToken)) return false;

        bool found = false;
        string? previous1 = null;
        int tokenCount = 0;
        foreach (string? token in EnumerateTokensCancellable(defaultAppearance!, cancellationToken)) {
            if (token == "Tf" && tokenCount >= 2 && previous1 is not null &&
                TryReadNumber(previous1, out double parsedFontSize) && parsedFontSize > 0D) {
                fontSize = parsedFontSize;
                found = true;
            }
            previous1 = token;
            tokenCount++;
        }
        return found;
    }

    private static System.Collections.Generic.IEnumerable<string?> EnumerateTokensCancellable(string source, CancellationToken cancellationToken) {
        for (int index = 0; index < source.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            while (index < source.Length && IsSeparator(source[index])) {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                index++;
            }
            int start = index;
            while (index < source.Length && !IsSeparator(source[index])) {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                index++;
            }
            if (start < index) {
                // PDF syntax already caps numeric tokens at 4096 characters. An
                // oversized appearance token cannot be parsed as a color or size.
                yield return index - start <= 4096
                    ? PdfEncoding.StringSliceCancellable(source, start, index - start, cancellationToken)
                    : null;
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static bool IsSeparator(char value) => value == ' ' || value == '\t' || value == '\r' || value == '\n';

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
