namespace OfficeIMO.Excel.Utilities;

internal static partial class ExcelSanitizer {
    /// <summary>Encodes semantic text as SpreadsheetML, preserving literal tokens and XML-illegal UTF-16 units.</summary>
    internal static string EncodeXString(string value) {
        int first = 0;
        while (first < value.Length && !IsXStringEscape(value, first) && !RequiresXStringEscape(value, first)) first++;
        if (first == value.Length) return value;
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
