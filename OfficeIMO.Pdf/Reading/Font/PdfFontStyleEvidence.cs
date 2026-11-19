namespace OfficeIMO.Pdf;

internal static class PdfFontStyleEvidence {
    private const int ItalicFlag = 1 << 6;

    internal static bool IsBold(string? baseFont, int? fontWeight) =>
        fontWeight.HasValue
            ? fontWeight.Value >= 700
            : HasStyleSuffix(baseFont, "Bold") ||
              HasStyleSuffix(baseFont, "Black") ||
              HasStyleSuffix(baseFont, "Heavy") ||
              HasStyleSuffix(baseFont, "Demi");

    internal static bool IsItalic(string? baseFont, int? fontDescriptorFlags) =>
        fontDescriptorFlags.HasValue
            ? (fontDescriptorFlags.Value & ItalicFlag) != 0
            : HasStyleSuffix(baseFont, "Italic") || HasStyleSuffix(baseFont, "Oblique");

    private static bool HasStyleSuffix(string? baseFont, string token) {
        if (string.IsNullOrWhiteSpace(baseFont)) return false;
        int subsetSeparator = baseFont!.Length > 7 && baseFont[6] == '+' ? 7 : 0;
        int end = WithoutIdentityCMapSuffix(baseFont);
        int styleSeparator = Math.Max(baseFont.LastIndexOf('-', end - 1), baseFont.LastIndexOf(',', end - 1));
        if (styleSeparator < subsetSeparator || styleSeparator + 1 >= end) return false;
        return baseFont.IndexOf(token, styleSeparator + 1, end - styleSeparator - 1, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    // A Type0 font's BaseFont is the descendant font name, a hyphen and the CMap name, as in "Carlito-Bold-Identity-H".
    private static int WithoutIdentityCMapSuffix(string baseFont) {
        foreach (string suffix in IdentityCMapSuffixes) {
            if (baseFont.Length > suffix.Length && baseFont.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) {
                return baseFont.Length - suffix.Length;
            }
        }

        return baseFont.Length;
    }

    private static readonly string[] IdentityCMapSuffixes = { "-Identity-H", "-Identity-V" };
}
