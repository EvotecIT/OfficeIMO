using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Maps common office font family names to dependency-free PDF standard fonts.
/// </summary>
public static class PdfStandardFontMapper {
    private static readonly char[] FamilySeparators = { ',', ';' };
    private static readonly char[] FamilyTrimChars = { ' ', '\t', '"', '\'' };

    /// <summary>
    /// Attempts to map an office font family name to the nearest PDF standard font family.
    /// </summary>
    public static bool TryMapFontFamily(string? fontFamily, out PdfStandardFont font) =>
        TryMapFontFamily(fontFamily, bold: false, italic: false, out font);

    /// <summary>
    /// Attempts to map an office font family name and style flags to the nearest PDF standard font variant.
    /// </summary>
    public static bool TryMapFontFamily(string? fontFamily, bool bold, bool italic, out PdfStandardFont font) {
        font = PdfStandardFont.Helvetica;
        if (string.IsNullOrWhiteSpace(fontFamily)) {
            return false;
        }

        foreach (string candidate in EnumerateNormalizedFamilies(fontFamily!)) {
            if (TryMapNormalizedFamily(candidate, out PdfStandardFont family)) {
                font = GetStyledFont(family, bold, italic);
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Returns the bold/italic variant for a supported PDF standard font family.
    /// </summary>
    public static PdfStandardFont GetStyledFont(PdfStandardFont font, bool bold, bool italic) {
        Guard.StandardFont(font, nameof(font), "PDF font must be one of the supported standard PDF fonts.");
        PdfStandardFont family = GetFontFamily(font);
        return family switch {
            PdfStandardFont.TimesRoman => bold && italic ? PdfStandardFont.TimesBoldItalic :
                bold ? PdfStandardFont.TimesBold :
                italic ? PdfStandardFont.TimesItalic :
                PdfStandardFont.TimesRoman,
            PdfStandardFont.Courier => bold && italic ? PdfStandardFont.CourierBoldOblique :
                bold ? PdfStandardFont.CourierBold :
                italic ? PdfStandardFont.CourierOblique :
                PdfStandardFont.Courier,
            _ => bold && italic ? PdfStandardFont.HelveticaBoldOblique :
                bold ? PdfStandardFont.HelveticaBold :
                italic ? PdfStandardFont.HelveticaOblique :
                PdfStandardFont.Helvetica
        };
    }

    /// <summary>
    /// Returns the normal family face for a supported PDF standard font variant.
    /// </summary>
    public static PdfStandardFont GetFontFamily(PdfStandardFont font) {
        Guard.StandardFont(font, nameof(font), "PDF font must be one of the supported standard PDF fonts.");
        return font switch {
            PdfStandardFont.TimesRoman or
            PdfStandardFont.TimesItalic or
            PdfStandardFont.TimesBold or
            PdfStandardFont.TimesBoldItalic => PdfStandardFont.TimesRoman,
            PdfStandardFont.Courier or
            PdfStandardFont.CourierOblique or
            PdfStandardFont.CourierBold or
            PdfStandardFont.CourierBoldOblique => PdfStandardFont.Courier,
            _ => PdfStandardFont.Helvetica
        };
    }

    /// <summary>
    /// Reports whether the first recognized family in an Office/CSS family list names the selected
    /// built-in PDF family directly rather than merely mapping to it as an approximation.
    /// </summary>
    public static bool IsStandardPdfFamilyEquivalent(string? fontFamily, PdfStandardFont font) {
        Guard.StandardFont(font, nameof(font), "PDF font must be one of the supported standard PDF fonts.");
        if (string.IsNullOrWhiteSpace(fontFamily)) {
            return false;
        }

        PdfStandardFont expectedFamily = GetFontFamily(font);
        foreach (string candidate in EnumerateNormalizedFamilies(fontFamily!)) {
            if (!TryMapNormalizedFamily(candidate, out PdfStandardFont mapped)) {
                continue;
            }

            if (GetFontFamily(mapped) != expectedFamily) {
                return false;
            }

            return expectedFamily switch {
                PdfStandardFont.TimesRoman => candidate is "times" or "timesroman" or "serif",
                PdfStandardFont.Courier => candidate is "courier" or "monospace",
                _ => candidate is "helvetica" or "sans" or "sansserif"
            };
        }

        return false;
    }

    private static IEnumerable<string> EnumerateNormalizedFamilies(string fontFamily) {
        foreach (string family in fontFamily.Split(FamilySeparators)) {
            string normalized = NormalizeSingleFontFamily(family);
            if (normalized.Length > 0) {
                yield return normalized;
            }
        }
    }

    private static string NormalizeSingleFontFamily(string fontFamily) {
        string firstFamily = fontFamily.Trim(FamilyTrimChars);
        var builder = new StringBuilder(firstFamily.Length);
        foreach (char ch in firstFamily) {
            if (char.IsLetterOrDigit(ch)) {
                builder.Append(char.ToLowerInvariant(ch));
            }
        }

        return builder.ToString();
    }

    private static bool TryMapNormalizedFamily(string normalized, out PdfStandardFont font) {
        switch (normalized) {
            case "timesnewroman":
            case "times":
            case "timesroman":
            case "georgia":
            case "cambria":
            case "serif":
                font = PdfStandardFont.TimesRoman;
                return true;
            case "couriernew":
            case "courier":
            case "consolas":
            case "lucidaconsole":
            case "monospace":
            case "monaco":
                font = PdfStandardFont.Courier;
                return true;
            case "arial":
            case "helvetica":
            case "calibri":
            case "aptos":
            case "segoeui":
            case "tahoma":
            case "verdana":
            case "sans":
            case "sansserif":
                font = PdfStandardFont.Helvetica;
                return true;
            default:
                font = PdfStandardFont.Helvetica;
                return false;
        }
    }
}
