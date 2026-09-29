namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    // Unicode 17.0 Scripts.txt, Script=Latin. Keep holes between assigned ranges:
    // https://www.unicode.org/Public/17.0.0/ucd/Scripts.txt
    private static bool IsLatinScriptScalar(int scalar) =>
        scalar >= 0x0041 && scalar <= 0x005A || scalar >= 0x0061 && scalar <= 0x007A ||
        scalar == 0x00AA || scalar == 0x00BA ||
        scalar >= 0x00C0 && scalar <= 0x00D6 || scalar >= 0x00D8 && scalar <= 0x00F6 ||
        scalar >= 0x00F8 && scalar <= 0x02B8 || scalar >= 0x02E0 && scalar <= 0x02E4 ||
        scalar >= 0x1D00 && scalar <= 0x1D25 || scalar >= 0x1D2C && scalar <= 0x1D5C ||
        scalar >= 0x1D62 && scalar <= 0x1D65 || scalar >= 0x1D6B && scalar <= 0x1D77 ||
        scalar >= 0x1D79 && scalar <= 0x1DBE || scalar >= 0x1E00 && scalar <= 0x1EFF ||
        scalar == 0x2071 || scalar == 0x207F || scalar >= 0x2090 && scalar <= 0x209C ||
        scalar >= 0x212A && scalar <= 0x212B || scalar == 0x2132 || scalar == 0x214E ||
        scalar >= 0x2160 && scalar <= 0x2188 || scalar >= 0x2C60 && scalar <= 0x2C7F ||
        scalar >= 0xA722 && scalar <= 0xA787 || scalar >= 0xA78B && scalar <= 0xA7DC ||
        scalar >= 0xA7F1 && scalar <= 0xA7FF || scalar >= 0xAB30 && scalar <= 0xAB5A ||
        scalar >= 0xAB5C && scalar <= 0xAB64 || scalar >= 0xAB66 && scalar <= 0xAB69 ||
        scalar >= 0xFB00 && scalar <= 0xFB06 || scalar >= 0xFF21 && scalar <= 0xFF3A ||
        scalar >= 0xFF41 && scalar <= 0xFF5A || scalar >= 0x10780 && scalar <= 0x10785 ||
        scalar >= 0x10787 && scalar <= 0x107B0 || scalar >= 0x107B2 && scalar <= 0x107BA ||
        scalar >= 0x1DF00 && scalar <= 0x1DF1E || scalar >= 0x1DF25 && scalar <= 0x1DF2A;

    // Common punctuation remains eligible only after contextual script resolution.
    private static bool IsLatinDefaultScalar(int scalar) => scalar <= 0x024F || IsLatinScriptScalar(scalar);
}
