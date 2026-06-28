namespace OfficeIMO.Markdown;

/// <summary>
/// Decodes CommonMark character references without depending only on the narrower
/// runtime HTML decoder table.
/// </summary>
internal static class CommonMarkCharacterReference {
    private const int MaxNamedReferenceLength = 32;
    private const int MaxDecimalReferenceDigits = 7;
    private const int MaxHexReferenceDigits = 8;

    private static readonly Dictionary<string, string> NamedReferences = new(StringComparer.Ordinal) {
        ["nbsp"] = "\u00A0",
        ["amp"] = "&",
        ["copy"] = "\u00A9",
        ["AElig"] = "\u00C6",
        ["Dcaron"] = "\u010E",
        ["frac34"] = "\u00BE",
        ["HilbertSpace"] = "\u210B",
        ["DifferentialD"] = "\u2146",
        ["ClockwiseContourIntegral"] = "\u2232",
        ["ngE"] = "\u2267\u0338"
    };

    private static readonly Dictionary<int, int> NumericReplacementMap = new() {
        [0x80] = 0x20AC,
        [0x82] = 0x201A,
        [0x83] = 0x0192,
        [0x84] = 0x201E,
        [0x85] = 0x2026,
        [0x86] = 0x2020,
        [0x87] = 0x2021,
        [0x88] = 0x02C6,
        [0x89] = 0x2030,
        [0x8A] = 0x0160,
        [0x8B] = 0x2039,
        [0x8C] = 0x0152,
        [0x8E] = 0x017D,
        [0x91] = 0x2018,
        [0x92] = 0x2019,
        [0x93] = 0x201C,
        [0x94] = 0x201D,
        [0x95] = 0x2022,
        [0x96] = 0x2013,
        [0x97] = 0x2014,
        [0x98] = 0x02DC,
        [0x99] = 0x2122,
        [0x9A] = 0x0161,
        [0x9B] = 0x203A,
        [0x9C] = 0x0153,
        [0x9E] = 0x017E,
        [0x9F] = 0x0178
    };

    internal static string DecodeAll(string value) {
        if (string.IsNullOrEmpty(value)) {
            return value;
        }

        StringBuilder? builder = null;
        int segmentStart = 0;

        for (int i = 0; i < value.Length; i++) {
            if (value[i] != '&' || !TryDecode(value, i, out int consumed, out string decoded)) {
                continue;
            }

            builder ??= new StringBuilder(value.Length);
            if (i > segmentStart) {
                builder.Append(value, segmentStart, i - segmentStart);
            }

            builder.Append(decoded);
            i += consumed - 1;
            segmentStart = i + 1;
        }

        if (builder == null) {
            return value;
        }

        if (segmentStart < value.Length) {
            builder.Append(value, segmentStart, value.Length - segmentStart);
        }

        return builder.ToString();
    }

    internal static bool TryDecode(string text, int start, out int consumed, out string decoded) {
        consumed = 0;
        decoded = string.Empty;

        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length || text[start] != '&') {
            return false;
        }

        if (start + 2 < text.Length && text[start + 1] == '#') {
            return TryDecodeNumeric(text, start, out consumed, out decoded);
        }

        return TryDecodeNamed(text, start, out consumed, out decoded);
    }

    private static bool TryDecodeNamed(string text, int start, out int consumed, out string decoded) {
        consumed = 0;
        decoded = string.Empty;

        int nameStart = start + 1;
        if (nameStart >= text.Length || !IsAsciiLetter(text[nameStart])) {
            return false;
        }

        int scan = nameStart + 1;
        while (scan < text.Length && scan - nameStart < MaxNamedReferenceLength && IsAsciiAlphanumeric(text[scan])) {
            scan++;
        }

        if (scan >= text.Length || text[scan] != ';') {
            return false;
        }

        string name = text.Substring(nameStart, scan - nameStart);
        if (NamedReferences.TryGetValue(name, out string? namedDecoded) && namedDecoded != null) {
            consumed = scan - start + 1;
            decoded = namedDecoded;
            return true;
        }

        string candidate = text.Substring(start, scan - start + 1);
        string? htmlDecoded = System.Net.WebUtility.HtmlDecode(candidate);
        if (string.IsNullOrEmpty(htmlDecoded) || string.Equals(htmlDecoded, candidate, StringComparison.Ordinal)) {
            return false;
        }

        consumed = candidate.Length;
        decoded = htmlDecoded;
        return true;
    }

    private static bool TryDecodeNumeric(string text, int start, out int consumed, out string decoded) {
        consumed = 0;
        decoded = string.Empty;

        int digitStart = start + 2;
        int numberBase = 10;
        int maxDigits = MaxDecimalReferenceDigits;
        if (digitStart < text.Length && (text[digitStart] == 'x' || text[digitStart] == 'X')) {
            numberBase = 16;
            maxDigits = MaxHexReferenceDigits;
            digitStart++;
        }

        if (digitStart >= text.Length) {
            return false;
        }

        int scan = digitStart;
        long value = 0;
        while (scan < text.Length && scan - digitStart < maxDigits) {
            int digit = DecodeDigit(text[scan], numberBase);
            if (digit < 0) {
                break;
            }

            value = (value * numberBase) + digit;
            if (value > 0x10FFFF) {
                value = 0x110000;
            }

            scan++;
        }

        if (scan == digitStart || scan >= text.Length || text[scan] != ';') {
            return false;
        }

        consumed = scan - start + 1;
        decoded = DecodeCodePoint(value);
        return true;
    }

    private static string DecodeCodePoint(long value) {
        if (value == 0 || value > 0x10FFFF || (value >= 0xD800 && value <= 0xDFFF)) {
            return "\uFFFD";
        }

        int codePoint = (int)value;
        if (NumericReplacementMap.TryGetValue(codePoint, out int replacement)) {
            codePoint = replacement;
        }

        return char.ConvertFromUtf32(codePoint);
    }

    private static int DecodeDigit(char value, int numberBase) {
        if (value >= '0' && value <= '9') {
            return value - '0';
        }

        if (numberBase != 16) {
            return -1;
        }

        if (value >= 'a' && value <= 'f') {
            return value - 'a' + 10;
        }

        if (value >= 'A' && value <= 'F') {
            return value - 'A' + 10;
        }

        return -1;
    }

    private static bool IsAsciiLetter(char value) =>
        (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');

    private static bool IsAsciiAlphanumeric(char value) =>
        IsAsciiLetter(value) || (value >= '0' && value <= '9');
}
