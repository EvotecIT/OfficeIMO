namespace OfficeIMO.Pdf;

/// <summary>
/// Maps PDF PostScript base-font names to the family names that office applications resolve, for editable
/// reconstruction. "ABCDEF+ArialMT" becomes "Arial", "TimesNewRomanPS-BoldMT" becomes "Times New Roman" and
/// "LiberationSans-Bold" becomes "Liberation Sans".
/// </summary>
internal static class PdfFontFamilyNames {
    private static readonly string[] StyleWords = {
        "BoldItalicMT", "BoldItalic", "BoldOblique", "BoldMT", "ItalicMT", "SemiBold", "Semibold", "DemiBold", "Demibold",
        "ExtraBold", "ExtraLight", "UltraBold", "Bold", "Italic", "Oblique", "Regular", "Roman", "Book", "Medium", "Light",
        "Black", "Heavy", "Thin", "Condensed", "Cond", "Narrow"
    };

    private static readonly string[] VendorSuffixes = { "MT", "PS" };

    private static readonly string[] StrippableSuffixes = new[] { "PSMT", "PS", "MT" }.Concat(StyleWords).ToArray();

    private static readonly Dictionary<string, string> KnownFamilies = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
        ["Arial"] = "Arial",
        ["ArialNarrow"] = "Arial Narrow",
        // Office does not install Helvetica; Arial shares its metrics and is what Office substitutes.
        ["Helvetica"] = "Arial",
        ["HelveticaNeue"] = "Helvetica Neue",
        ["Times"] = "Times New Roman",
        ["TimesNewRoman"] = "Times New Roman",
        ["Courier"] = "Courier New",
        ["CourierNew"] = "Courier New",
        ["Symbol"] = "Symbol",
        ["ZapfDingbats"] = "Zapf Dingbats",
        ["SegoeUI"] = "Segoe UI",
        ["CalibriLight"] = "Calibri Light",
        ["CambriaMath"] = "Cambria Math",
        ["TrebuchetMS"] = "Trebuchet MS",
        ["ComicSansMS"] = "Comic Sans MS",
        ["LucidaConsole"] = "Lucida Console",
        ["LucidaGrande"] = "Lucida Grande",
        ["SFPro"] = "SF Pro",
        ["SFProText"] = "SF Pro Text",
        ["SFProDisplay"] = "SF Pro Display",
        [".SFNS"] = "SF Pro",
        ["MSGothic"] = "MS Gothic",
        ["MSMincho"] = "MS Mincho",
        ["MSPGothic"] = "MS PGothic",
        ["DejaVuSans"] = "DejaVu Sans",
        ["DejaVuSerif"] = "DejaVu Serif",
        ["DejaVuSansMono"] = "DejaVu Sans Mono",
        ["TimesNewRomanPS"] = "Times New Roman"
    };

    internal static string? ToFamilyName(string? baseFont) {
        if (string.IsNullOrWhiteSpace(baseFont)) return null;
        string value = baseFont!.Trim();
        if (value.Length > 7 && value[6] == '+' && value.Take(6).All(static ch => ch >= 'A' && ch <= 'Z')) {
            value = value.Substring(7);
        }
        value = value.Replace('_', ' ').Replace('#', ' ');
        // "Arial,Bold" and "Arial-BoldMT": the part after the separator names the style when it is a style word.
        int separator = -1;
        for (int index = 0; index < value.Length && separator < 0; index++) {
            if (value[index] == ',' || value[index] == '-') separator = index;
        }
        if (separator > 0) {
            string head = value.Substring(0, separator);
            string tail = value.Substring(separator + 1);
            value = IsStyleSuffix(tail) || head.Length > 3 ? head : head + tail;
        }
        value = value.Trim();
        for (bool changed = true; changed;) {
            changed = false;
            if (KnownFamilies.TryGetValue(value.Replace(" ", string.Empty), out string? knownBeforeStrip)) return knownBeforeStrip;
            foreach (string suffix in StrippableSuffixes) {
                if (value.Length > suffix.Length + 1 && value.EndsWith(suffix, StringComparison.Ordinal)) {
                    value = value.Substring(0, value.Length - suffix.Length).TrimEnd();
                    changed = true;
                    break;
                }
            }
        }
        if (value.Length == 0) return null;
        string compact = value.Replace(" ", string.Empty);
        if (KnownFamilies.TryGetValue(compact, out string? known)) return known;
        return value.Any(static ch => ch == ' ') ? value : SplitCamelCase(value);
    }

    private static bool IsStyleSuffix(string value) {
        string remaining = value.Trim();
        if (remaining.Length == 0) return true;
        foreach (string suffix in VendorSuffixes) {
            if (remaining.EndsWith(suffix, StringComparison.Ordinal)) remaining = remaining.Substring(0, remaining.Length - suffix.Length);
        }
        while (remaining.Length > 0) {
            string? match = StyleWords.FirstOrDefault(word => remaining.StartsWith(word, StringComparison.OrdinalIgnoreCase));
            if (match == null) return false;
            remaining = remaining.Substring(match.Length);
        }
        return true;
    }

    private static string SplitCamelCase(string value) {
        var builder = new System.Text.StringBuilder(value.Length + 4);
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (index > 0 && char.IsUpper(current) &&
                (char.IsLower(value[index - 1]) || index + 1 < value.Length && char.IsLower(value[index + 1]) && char.IsUpper(value[index - 1]))) {
                builder.Append(' ');
            }
            builder.Append(current);
        }
        return builder.ToString();
    }
}
