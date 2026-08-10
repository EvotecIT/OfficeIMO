namespace OfficeIMO.Pdf;

/// <summary>PDFDocEncoding conversion for PDF text strings without a Unicode byte-order marker.</summary>
internal static class PdfDocEncoding {
    private static readonly char[] Map = BuildMap();

    internal static bool TryDecode(byte[] bytes, out string value) {
        if (bytes.Length == 0) {
            value = string.Empty;
            return true;
        }

        var characters = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) {
            char character = Map[bytes[i]];
            if (character == '\0') {
                value = string.Empty;
                return false;
            }
            characters[i] = character;
        }
        value = new string(characters);
        return true;
    }

    private static char[] BuildMap() {
        var map = new char[256];
        map[9] = '\t'; map[10] = '\n'; map[13] = '\r';
        map[24] = '\u02D8'; map[25] = '\u02C7'; map[26] = '\u02C6'; map[27] = '\u02D9';
        map[28] = '\u02DD'; map[29] = '\u02DB'; map[30] = '\u02DA'; map[31] = '\u02DC';
        for (int value = 32; value <= 126; value++) map[value] = (char)value;
        char[] specials = {
            '\u2022', '\u2020', '\u2021', '\u2026', '\u2014', '\u2013', '\u0192', '\u2044',
            '\u2039', '\u203A', '\u2212', '\u2030', '\u201E', '\u201C', '\u201D', '\u2018',
            '\u2019', '\u201A', '\u2122', '\uFB01', '\uFB02', '\u0141', '\u0152', '\u0160',
            '\u0178', '\u017D', '\u0131', '\u0142', '\u0153', '\u0161', '\u017E'
        };
        for (int index = 0; index < specials.Length; index++) map[128 + index] = specials[index];
        map[160] = '\u20AC';
        for (int value = 161; value <= 172; value++) map[value] = (char)value;
        for (int value = 174; value <= 255; value++) map[value] = (char)value;
        return map;
    }
}
