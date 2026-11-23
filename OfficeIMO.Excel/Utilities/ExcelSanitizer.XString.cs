namespace OfficeIMO.Excel.Utilities;

internal static partial class ExcelSanitizer {
#if NET8_0_OR_GREATER
    // Characters that can start an XString escape or need one: the underscore of a literal _xHHHH_ token, and
    // controls other than tab, line feed and carriage return. Surrogates and U+FFFE/U+FFFF are searched separately.
    private static readonly System.Buffers.SearchValues<char> XStringAsciiCandidates =
        System.Buffers.SearchValues.Create("\u0000\u0001\u0002\u0003\u0004\u0005\u0006\u0007\u0008\u000B\u000C\u000E\u000F\u0010\u0011\u0012\u0013\u0014\u0015\u0016\u0017\u0018\u0019\u001A\u001B\u001C\u001D\u001E\u001F_");
#endif

    /// <summary>Encodes semantic text as SpreadsheetML, preserving literal tokens and XML-illegal UTF-16 units.</summary>
    internal static string EncodeXString(string value) {
        int first = FindXStringEncodingStart(value);
        if (first < 0) return value;
        var builder = new StringBuilder(value.Length + 6);
        builder.Append(value, 0, first);
        for (int index = first; index < value.Length; index++) {
            if (IsXStringEscape(value, index)) builder.Append("_x005F_");
            else if (RequiresXStringEscape(value, index)) {
                builder.Append("_x");
                builder.Append(((int)value[index]).ToString("X4", System.Globalization.CultureInfo.InvariantCulture));
                builder.Append('_');
            } else builder.Append(value[index]);
        }
        return builder.ToString();
    }

    /// <summary>Returns the first index that <see cref="EncodeXString"/> must rewrite, or -1.</summary>
    private static int FindXStringEncodingStart(string value) {
#if NET8_0_OR_GREATER
        // Most text is BMP characters from the space up with no underscore, and needs no further checks.
        if (value.AsSpan().IndexOfAnyExceptInRange(' ', '\uD7FF') < 0 && value.IndexOf('_') < 0) return -1;
        // Line breaks and other text outside the fast range must not push long strings onto a per-character scan.
        // Each kind of candidate keeps its next position until the scan passes it, so runs of valid surrogate pairs
        // do not search the rest of the string again for every pair.
        int ascii = -2, surrogate = -2, nonCharacter = -2;
        int start = 0;
        while (start < value.Length) {
            ReadOnlySpan<char> remaining = value.AsSpan(start);
            if (ascii != -1 && ascii < start) ascii = Offset(start, remaining.IndexOfAny(XStringAsciiCandidates));
            if (surrogate != -1 && surrogate < start) surrogate = Offset(start, remaining.IndexOfAnyInRange('\uD800', '\uDFFF'));
            if (nonCharacter != -1 && nonCharacter < start) nonCharacter = Offset(start, remaining.IndexOfAny('\uFFFE', '\uFFFF'));
            int index = Earliest(Earliest(ascii, surrogate), nonCharacter);
            if (index < 0) return -1;
            if (IsXStringEscape(value, index) || RequiresXStringEscape(value, index)) return index;
            start = index + 1;
        }
        return -1;

        static int Offset(int start, int found) => found < 0 ? -1 : start + found;
        static int Earliest(int left, int right) => left < 0 ? right : right < 0 ? left : Math.Min(left, right);
#else
        for (int index = 0; index < value.Length; index++) {
            if (IsXStringEscape(value, index) || RequiresXStringEscape(value, index)) return index;
        }
        return -1;
#endif
    }

    /// <summary>Decodes an XML string payload once; an escaped underscore cannot start a second decoding pass.</summary>
    internal static string DecodeXString(string value) {
        int first = FindXStringEscape(value);
        if (first < 0) return value;
        var builder = new StringBuilder(value.Length);
        builder.Append(value, 0, first);
        for (int index = first; index < value.Length; index++) {
            if (IsXStringEscape(value, index)) {
                int code = 0;
                for (int digit = index + 2; digit < index + 6; digit++) code = (code << 4) | HexValue(value[digit]);
                builder.Append((char)code);
                index += 6;
            } else builder.Append(value[index]);
        }
        return builder.ToString();
    }

    internal static bool HasXStringEscape(ReadOnlySpan<byte> value) {
        for (int index = value.IndexOf((byte)'_'); index >= 0 && index + 6 < value.Length; index++) {
            if (value[index] == (byte)'_' && value[index + 1] == (byte)'x' && value[index + 6] == (byte)'_'
                && HexValue((char)value[index + 2]) >= 0 && HexValue((char)value[index + 3]) >= 0
                && HexValue((char)value[index + 4]) >= 0 && HexValue((char)value[index + 5]) >= 0) return true;
        }
        return false;
    }

    private static int FindXStringEscape(string value) {
        for (int index = value.IndexOf('_'); index >= 0 && index + 6 < value.Length; index++)
            if (IsXStringEscape(value, index)) return index;
        return -1;
    }

    private static bool IsXStringEscape(string value, int index) => index + 6 < value.Length
        && value[index] == '_' && value[index + 1] == 'x' && value[index + 6] == '_'
        && HexValue(value[index + 2]) >= 0 && HexValue(value[index + 3]) >= 0
        && HexValue(value[index + 4]) >= 0 && HexValue(value[index + 5]) >= 0;

    private static bool RequiresXStringEscape(string value, int index) {
        char current = value[index];
        if (current < ' ') return current != '\t' && current != '\n' && current != '\r';
        if (current == '\uFFFE' || current == '\uFFFF') return true;
        if (char.IsHighSurrogate(current)) return index + 1 == value.Length || !char.IsLowSurrogate(value[index + 1]);
        if (char.IsLowSurrogate(current)) return index == 0 || !char.IsHighSurrogate(value[index - 1]);
        return false;
    }

    private static int HexValue(char value) => value >= '0' && value <= '9' ? value - '0'
        : value >= 'a' && value <= 'f' ? value - 'a' + 10
        : value >= 'A' && value <= 'F' ? value - 'A' + 10 : -1;
}
