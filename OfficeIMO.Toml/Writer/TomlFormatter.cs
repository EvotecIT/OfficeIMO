namespace OfficeIMO.Toml;

/// <summary>
/// Writes keys and values as TOML text. <see cref="TomlWriter"/> and document editing share it, so a value is spelled
/// the same way whichever API writes it. Output stays readable by TOML 1.0 parsers: no <c>\e</c> or <c>\x</c> escapes and
/// seconds are always written.
/// </summary>
internal static class TomlFormatter {
    public static bool IsBareKey(string key) {
        if (key.Length == 0) return false;
        for (int i = 0; i < key.Length; i++) {
            if (!TomlScalars.IsBareKeyChar(key[i])) return false;
        }
        return true;
    }

    public static void AppendKey(StringBuilder text, string key) {
        if (IsBareKey(key)) text.Append(key);
        else AppendString(text, key, TomlStringStyle.Basic);
    }

    public static void AppendKeys(StringBuilder text, IReadOnlyList<string> keys) {
        for (int i = 0; i < keys.Count; i++) {
            if (i > 0) text.Append('.');
            AppendKey(text, keys[i]);
        }
    }

    /// <summary>Returns <paramref name="preferred"/> when it can hold the value exactly, otherwise a basic string.</summary>
    public static TomlStringStyle FitStyle(string value, TomlStringStyle preferred) {
        switch (preferred) {
            case TomlStringStyle.Literal:
                for (int i = 0; i < value.Length; i++) {
                    char ch = value[i];
                    if (ch == '\'' || ch == '\n' || ch == '\r' || TomlScalars.IsForbiddenControl(ch)) return TomlStringStyle.Basic;
                }
                return TomlStringStyle.Literal;
            case TomlStringStyle.MultilineLiteral:
                if (value.IndexOf("'''", StringComparison.Ordinal) >= 0) return TomlStringStyle.MultilineBasic;
                for (int i = 0; i < value.Length; i++) {
                    char ch = value[i];
                    if (ch == '\r' || (ch != '\n' && TomlScalars.IsForbiddenControl(ch))) return TomlStringStyle.MultilineBasic;
                }
                return TomlStringStyle.MultilineLiteral;
            case TomlStringStyle.MultilineBasic:
                return TomlStringStyle.MultilineBasic;
            default:
                return TomlStringStyle.Basic;
        }
    }

    public static void AppendString(StringBuilder text, string value, TomlStringStyle style) {
        switch (style) {
            case TomlStringStyle.Literal:
                text.Append('\'').Append(value).Append('\'');
                return;
            case TomlStringStyle.MultilineLiteral:
                // The newline after the opening delimiter is trimmed by readers, so a value starting with one survives.
                text.Append("'''\n").Append(value).Append("'''");
                return;
            case TomlStringStyle.MultilineBasic:
                text.Append("\"\"\"\n");
                AppendEscaped(text, value, multiline: true);
                text.Append("\"\"\"");
                return;
            default:
                text.Append('"');
                AppendEscaped(text, value, multiline: false);
                text.Append('"');
                return;
        }
    }

    private static void AppendEscaped(StringBuilder text, string value, bool multiline) {
        int quotes = 0;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch == '"') {
                // In a multi-line string every third quote in a row is escaped so no run closes the string.
                if (!multiline || ++quotes == 3) {
                    text.Append("\\\"");
                    quotes = 0;
                } else {
                    text.Append('"');
                }
                continue;
            }
            quotes = 0;
            switch (ch) {
                case '\\': text.Append("\\\\"); break;
                case '\n': text.Append(multiline ? "\n" : "\\n"); break;
                case '\t': text.Append("\\t"); break;
                case '\r': text.Append("\\r"); break;
                case '\b': text.Append("\\b"); break;
                case '\f': text.Append("\\f"); break;
                default:
                    if (TomlScalars.IsForbiddenControl(ch)) {
                        text.Append("\\u").Append(((int)ch).ToString("X4", CultureInfo.InvariantCulture));
                    } else {
                        text.Append(ch);
                    }
                    break;
            }
        }
        // A quote right before the closing delimiter would merge with it beyond the two quotes TOML allows.
        if (multiline && quotes > 0) {
            text.Length -= quotes;
            for (int i = 0; i < quotes; i++) text.Append("\\\"");
        }
    }

    public static void AppendInteger(StringBuilder text, long value, TomlIntegerBase numberBase, bool upperHex) {
        if (value < 0 || numberBase == TomlIntegerBase.Decimal) {
            text.Append(value.ToString(CultureInfo.InvariantCulture));
            return;
        }
        switch (numberBase) {
            case TomlIntegerBase.Hexadecimal:
                text.Append("0x").Append(value.ToString(upperHex ? "X" : "x", CultureInfo.InvariantCulture));
                break;
            case TomlIntegerBase.Octal:
                text.Append("0o").Append(Convert.ToString(value, 8));
                break;
            default:
                text.Append("0b").Append(Convert.ToString(value, 2));
                break;
        }
    }

    public static void AppendFloat(StringBuilder text, double value) {
        if (double.IsNaN(value)) {
            text.Append("nan");
            return;
        }
        if (double.IsPositiveInfinity(value)) {
            text.Append("inf");
            return;
        }
        if (double.IsNegativeInfinity(value)) {
            text.Append("-inf");
            return;
        }
        if (value == 0) {
            text.Append(BitConverter.DoubleToInt64Bits(value) < 0 ? "-0.0" : "0.0");
            return;
        }
        string number = value.ToString("R", CultureInfo.InvariantCulture);
        int exponent = number.IndexOf('E');
        if (exponent >= 0) {
            // "1E+20" is valid TOML; a mantissa without a fraction is fine before an exponent.
            text.Append(number, 0, exponent).Append('e').Append(number, exponent + 1, number.Length - exponent - 1);
            return;
        }
        text.Append(number);
        if (number.IndexOf('.') < 0) text.Append(".0");
    }

    public static void AppendDateTime(StringBuilder text, TomlDateTime value, char separator) {
        if (value.Kind != TomlDateTimeKind.LocalTime) {
            Pad(text, value.Year, 4).Append('-');
            Pad(text, value.Month, 2).Append('-');
            Pad(text, value.Day, 2);
            if (value.Kind == TomlDateTimeKind.LocalDate) return;
            text.Append(separator);
        }
        Pad(text, value.Hour, 2).Append(':');
        Pad(text, value.Minute, 2).Append(':');
        Pad(text, value.Second, 2);
        if (value.FractionDigits > 0) {
            int digits = value.FractionDigits;
            int fraction = value.Nanosecond;
            for (int i = digits; i < 9; i++) fraction /= 10;
            text.Append('.');
            Pad(text, fraction, digits);
        }
        if (value.Kind != TomlDateTimeKind.OffsetDateTime) return;
        int offset = value.OffsetMinutes;
        if (offset == 0) {
            text.Append('Z');
            return;
        }
        text.Append(offset < 0 ? '-' : '+');
        offset = Math.Abs(offset);
        Pad(text, offset / 60, 2).Append(':');
        Pad(text, offset % 60, 2);
    }

    private static StringBuilder Pad(StringBuilder text, int value, int width) {
        string digits = value.ToString(CultureInfo.InvariantCulture);
        for (int i = digits.Length; i < width; i++) text.Append('0');
        return text.Append(digits);
    }

    /// <summary>Writes a value node's current value, keeping the conventions of its source spelling.</summary>
    public static void AppendScalar(StringBuilder text, TomlValue node) {
        object value = node.Value;
        switch (node._kind) {
            case TomlNodeKind.String:
                AppendString(text, (string)value, node._style == TomlStringStyle.Bare ? TomlStringStyle.Basic : node._style);
                return;
            case TomlNodeKind.Integer:
                AppendInteger(text, (long)value, node._base, upperHex: node.IsSource && HasUpperHex(node));
                return;
            case TomlNodeKind.Float:
                AppendFloat(text, (double)value);
                return;
            case TomlNodeKind.Boolean:
                text.Append((bool)value ? "true" : "false");
                return;
            default:
                AppendDateTime(text, (TomlDateTime)value, SeparatorOf(node));
                return;
        }
    }

    private static bool HasUpperHex(TomlValue node) {
        string source = node._document._source;
        for (int i = node._start; i < node._end; i++) {
            if (source[i] >= 'A' && source[i] <= 'F') return true;
        }
        return false;
    }

    /// <summary>Returns the date/time separator of a loaded date-time (<c>T</c>, <c>t</c> or a space), or <c>T</c>.</summary>
    private static char SeparatorOf(TomlValue node) {
        if (!node.IsSource || node._end - node._start <= 10) return 'T';
        char ch = node._document._source[node._start + 10];
        return ch == ' ' || ch == 't' ? ch : 'T';
    }
}
