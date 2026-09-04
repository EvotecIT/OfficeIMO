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
        int styleSeparator = Math.Max(baseFont.LastIndexOf('-'), baseFont.LastIndexOf(','));
        if (styleSeparator < subsetSeparator || styleSeparator + 1 >= baseFont.Length) return false;
        return baseFont.IndexOf(token, styleSeparator + 1, StringComparison.OrdinalIgnoreCase) >= 0;
    }
}
