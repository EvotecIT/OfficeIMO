namespace OfficeIMO.Toml;

/// <summary>How an integer is written, kept so an edit can keep the base.</summary>
internal enum TomlIntegerBase : byte {
    Decimal = 0,
    Hexadecimal,
    Octal,
    Binary
}

/// <summary>
/// Character classes, validation and decoding for TOML keys and values. Validation runs once while scanning and
/// allocates nothing; decoding runs lazily when a value is read.
/// </summary>
internal static class TomlScalars {
    public static bool IsBareKeyChar(char ch) =>
        (uint)((ch | 0x20) - 'a') <= 'z' - 'a' || (uint)(ch - '0') <= 9 || ch == '_' || ch == '-';

    /// <summary>Characters that can appear in an unquoted value: numbers, booleans, special floats and date-times.</summary>
    public static bool IsAtomChar(char ch) =>
        (uint)((ch | 0x20) - 'a') <= 'z' - 'a' || (uint)(ch - '0') <= 9 || ch == '_' || ch == '-' || ch == '+' || ch == '.' || ch == ':';

    public static bool IsDigit(char ch) => (uint)(ch - '0') <= 9;

    public static int HexValue(char ch) {
        if ((uint)(ch - '0') <= 9) return ch - '0';
        int lower = ch | 0x20;
        if ((uint)(lower - 'a') <= 'f' - 'a') return lower - 'a' + 10;
        return -1;
    }

    /// <summary>Characters TOML forbids in strings and comments: controls other than tab, and DEL.</summary>
    public static bool IsForbiddenControl(char ch) => (ch < 0x20 && ch != '\t') || ch == 0x7F;

    // ---- Classification -------------------------------------------------------------------------------------------

    /// <summary>
    /// Classifies and fully validates an unquoted value. Returns the token type, or <see cref="TomlTokenType.None"/> with
    /// an error code and the index of the offending character.
    /// </summary>
    public static TomlTokenType Classify(ReadOnlySpan<char> atom, out string? error, out int errorIndex) {
        error = null;
        errorIndex = 0;
        if (atom.SequenceEqual("true".AsSpan()) || atom.SequenceEqual("false".AsSpan())) return TomlTokenType.Boolean;
        if (atom.Length >= 3 && IsDigit(atom[0]) && IsDigit(atom[1])) {
            if (atom[2] == ':' || (atom.Length >= 5 && IsDigit(atom[2]) && IsDigit(atom[3]) && atom[4] == '-')) {
                error = TryParseDateTime(atom, out TomlDateTime value, out errorIndex);
                if (error != null) return TomlTokenType.None;
                return value.Kind switch {
                    TomlDateTimeKind.OffsetDateTime => TomlTokenType.OffsetDateTime,
                    TomlDateTimeKind.LocalDateTime => TomlTokenType.LocalDateTime,
                    TomlDateTimeKind.LocalDate => TomlTokenType.LocalDate,
                    _ => TomlTokenType.LocalTime
                };
            }
        }
        int start = atom.Length > 0 && (atom[0] == '+' || atom[0] == '-') ? 1 : 0;
        ReadOnlySpan<char> body = atom.Slice(start);
        if (body.SequenceEqual("inf".AsSpan()) || body.SequenceEqual("nan".AsSpan())) return TomlTokenType.Float;
        if (body.Length >= 2 && body[0] == '0' && (body[1] == 'x' || body[1] == 'o' || body[1] == 'b')) {
            error = TryParseInteger(atom, out _, out _, out errorIndex);
            return error == null ? TomlTokenType.Integer : TomlTokenType.None;
        }
        if (body.IndexOfAny('.', 'e', 'E') >= 0) {
            error = ValidateFloat(atom, out errorIndex);
            return error == null ? TomlTokenType.Float : TomlTokenType.None;
        }
        error = TryParseInteger(atom, out _, out _, out errorIndex);
        return error == null ? TomlTokenType.Integer : TomlTokenType.None;
    }

    // ---- Integers -------------------------------------------------------------------------------------------------

    /// <summary>Parses a validated or unvalidated integer. Returns null on success, otherwise an error code.</summary>
    public static string? TryParseInteger(ReadOnlySpan<char> s, out long value, out TomlIntegerBase numberBase, out int errorIndex) {
        value = 0;
        numberBase = TomlIntegerBase.Decimal;
        errorIndex = 0;
        if (s.Length == 0) return TomlErrorCodes.InvalidInteger;
        int i = 0;
        bool negative = false;
        bool signed = false;
        if (s[0] == '+' || s[0] == '-') {
            negative = s[0] == '-';
            signed = true;
            i = 1;
        }
        int radix = 10;
        if (s.Length - i >= 2 && s[i] == '0' && (s[i + 1] == 'x' || s[i + 1] == 'o' || s[i + 1] == 'b')) {
            if (signed) return TomlErrorCodes.InvalidInteger;
            numberBase = s[i + 1] == 'x' ? TomlIntegerBase.Hexadecimal : s[i + 1] == 'o' ? TomlIntegerBase.Octal : TomlIntegerBase.Binary;
            radix = s[i + 1] == 'x' ? 16 : s[i + 1] == 'o' ? 8 : 2;
            i += 2;
        } else if (s.Length - i > 1 && s[i] == '0') {
            errorIndex = i;
            return TomlErrorCodes.InvalidInteger; // leading zero
        }
        ulong limit = negative ? 9223372036854775808UL : 9223372036854775807UL;
        ulong magnitude = 0;
        bool previousDigit = false;
        int digits = 0;
        for (; i < s.Length; i++) {
            char ch = s[i];
            if (ch == '_') {
                if (!previousDigit || i + 1 >= s.Length) {
                    errorIndex = i;
                    return TomlErrorCodes.InvalidInteger;
                }
                previousDigit = false;
                continue;
            }
            int digit = HexValue(ch);
            if (digit < 0 || digit >= radix) {
                errorIndex = i;
                return TomlErrorCodes.InvalidInteger;
            }
            ulong next = magnitude * (ulong)radix + (ulong)digit;
            if (magnitude > (ulong.MaxValue - (ulong)digit) / (ulong)radix || next > limit) {
                errorIndex = i;
                return TomlErrorCodes.InvalidInteger;
            }
            magnitude = next;
            previousDigit = true;
            digits++;
        }
        if (digits == 0 || !previousDigit) {
            errorIndex = s.Length;
            return TomlErrorCodes.InvalidInteger;
        }
        value = negative ? (long)(0UL - magnitude) : (long)magnitude;
        return null;
    }

    // ---- Floats ---------------------------------------------------------------------------------------------------

    public static string? ValidateFloat(ReadOnlySpan<char> s, out int errorIndex) {
        errorIndex = 0;
        int i = 0;
        if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
        int body = i;
        if (!ScanDigits(s, ref i, allowLeadingZeros: false)) {
            errorIndex = i;
            return TomlErrorCodes.InvalidFloat;
        }
        bool fraction = false, exponent = false;
        if (i < s.Length && s[i] == '.') {
            i++;
            if (!ScanDigits(s, ref i, allowLeadingZeros: true)) {
                errorIndex = i;
                return TomlErrorCodes.InvalidFloat;
            }
            fraction = true;
        }
        if (i < s.Length && (s[i] == 'e' || s[i] == 'E')) {
            i++;
            if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
            if (!ScanDigits(s, ref i, allowLeadingZeros: true)) {
                errorIndex = i;
                return TomlErrorCodes.InvalidFloat;
            }
            exponent = true;
        }
        if (i != s.Length || (!fraction && !exponent) || i == body) {
            errorIndex = i;
            return TomlErrorCodes.InvalidFloat;
        }
        return null;
    }

    /// <summary>Scans digits with single underscores between them.</summary>
    private static bool ScanDigits(ReadOnlySpan<char> s, ref int i, bool allowLeadingZeros) {
        int start = i;
        bool previousDigit = false;
        while (i < s.Length) {
            char ch = s[i];
            if (IsDigit(ch)) {
                previousDigit = true;
                i++;
            } else if (ch == '_' && previousDigit && i + 1 < s.Length && IsDigit(s[i + 1])) {
                previousDigit = false;
                i++;
            } else {
                break;
            }
        }
        if (i == start || !previousDigit) return false;
        if (!allowLeadingZeros && s[start] == '0' && i - start > 1) {
            i = start + 1;
            return false;
        }
        return true;
    }

    /// <summary>Parses a validated float.</summary>
    public static double ParseFloat(ReadOnlySpan<char> s) {
        int start = s.Length > 0 && (s[0] == '+' || s[0] == '-') ? 1 : 0;
        bool negative = s.Length > 0 && s[0] == '-';
        ReadOnlySpan<char> body = s.Slice(start);
        if (body.SequenceEqual("inf".AsSpan())) return negative ? double.NegativeInfinity : double.PositiveInfinity;
        if (body.SequenceEqual("nan".AsSpan())) return double.NaN;
        Span<char> buffer = s.Length <= 128 ? stackalloc char[s.Length] : new char[s.Length];
        int length = 0;
        for (int i = 0; i < s.Length; i++) {
            if (s[i] != '_') buffer[length++] = s[i];
        }
#if NETSTANDARD2_0 || NET472
        try {
            return double.Parse(buffer.Slice(0, length).ToString(), NumberStyles.Float, CultureInfo.InvariantCulture);
        } catch (OverflowException) {
            // .NET Framework throws where modern .NET rounds to infinity, as IEEE 754 does.
            return negative ? double.NegativeInfinity : double.PositiveInfinity;
        }
#else
        return double.Parse(buffer.Slice(0, length), NumberStyles.Float, CultureInfo.InvariantCulture);
#endif
    }

    // ---- Date-times -----------------------------------------------------------------------------------------------

    /// <summary>Parses and validates a date, time or date-time. Returns null on success, otherwise an error code.</summary>
    public static string? TryParseDateTime(ReadOnlySpan<char> s, out TomlDateTime value, out int errorIndex) {
        value = default;
        errorIndex = 0;
        int i = 0;
        if (s.Length >= 3 && s[2] == ':') {
            if (!TryParseTime(s, ref i, out int h, out int m, out int sec, out int ns, out int digits) || i != s.Length) {
                errorIndex = i;
                return TomlErrorCodes.InvalidDateTime;
            }
            value = TomlDateTime.Create(TomlDateTimeKind.LocalTime, 0, 0, 0, h, m, sec, ns, digits, 0);
            return null;
        }
        if (s.Length < 10 || !Digits(s, 0, 4) || s[4] != '-' || !Digits(s, 5, 2) || s[7] != '-' || !Digits(s, 8, 2)) {
            errorIndex = Math.Min(s.Length, 10);
            return TomlErrorCodes.InvalidDateTime;
        }
        int year = Number(s, 0, 4), month = Number(s, 5, 2), day = Number(s, 8, 2);
        if (!IsValidDate(year, month, day)) {
            errorIndex = 5;
            return TomlErrorCodes.InvalidDateTime;
        }
        i = 10;
        if (i == s.Length) {
            value = TomlDateTime.Create(TomlDateTimeKind.LocalDate, year, month, day, 0, 0, 0, 0, 0, 0);
            return null;
        }
        if (s[i] != 'T' && s[i] != 't' && s[i] != ' ') {
            errorIndex = i;
            return TomlErrorCodes.InvalidDateTime;
        }
        i++;
        if (!TryParseTime(s, ref i, out int hour, out int minute, out int second, out int nanos, out int fraction)) {
            errorIndex = i;
            return TomlErrorCodes.InvalidDateTime;
        }
        if (i == s.Length) {
            value = TomlDateTime.Create(TomlDateTimeKind.LocalDateTime, year, month, day, hour, minute, second, nanos, fraction, 0);
            return null;
        }
        int offset;
        if (s[i] == 'Z' || s[i] == 'z') {
            offset = 0;
            i++;
        } else if ((s[i] == '+' || s[i] == '-') && s.Length - i >= 6 && Digits(s, i + 1, 2) && s[i + 3] == ':' && Digits(s, i + 4, 2)) {
            int oh = Number(s, i + 1, 2), om = Number(s, i + 4, 2);
            if (oh > 23 || om > 59) {
                errorIndex = i;
                return TomlErrorCodes.InvalidDateTime;
            }
            offset = (oh * 60 + om) * (s[i] == '-' ? -1 : 1);
            i += 6;
        } else {
            errorIndex = i;
            return TomlErrorCodes.InvalidDateTime;
        }
        if (i != s.Length) {
            errorIndex = i;
            return TomlErrorCodes.InvalidDateTime;
        }
        value = TomlDateTime.Create(TomlDateTimeKind.OffsetDateTime, year, month, day, hour, minute, second, nanos, fraction, offset);
        return null;
    }

    private static bool IsValidDate(int year, int month, int day) {
        if (month < 1 || month > 12 || day < 1) return false;
        int max = month == 2 ? (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) ? 29 : 28)
            : month == 4 || month == 6 || month == 9 || month == 11 ? 30 : 31;
        return day <= max;
    }

    /// <summary>Parses <c>HH:MM[:SS[.fraction]]</c>; TOML 1.1 makes the seconds optional.</summary>
    private static bool TryParseTime(ReadOnlySpan<char> s, ref int i, out int hour, out int minute, out int second, out int nanos, out int digits) {
        hour = minute = second = nanos = digits = 0;
        if (s.Length - i < 5 || !Digits(s, i, 2) || s[i + 2] != ':' || !Digits(s, i + 3, 2)) return false;
        hour = Number(s, i, 2);
        minute = Number(s, i + 3, 2);
        if (hour > 23 || minute > 59) return false;
        i += 5;
        if (i < s.Length && s[i] == ':') {
            if (s.Length - i < 3 || !Digits(s, i + 1, 2)) return false;
            second = Number(s, i + 1, 2);
            if (second > 60) return false;
            i += 3;
            if (i < s.Length && s[i] == '.') {
                i++;
                int start = i;
                int scale = 100_000_000;
                while (i < s.Length && IsDigit(s[i])) {
                    if (scale > 0) {
                        nanos += (s[i] - '0') * scale;
                        scale /= 10;
                    }
                    i++;
                }
                if (i == start) return false;
                digits = Math.Min(9, i - start);
            }
        }
        return true;
    }

    private static bool Digits(ReadOnlySpan<char> s, int start, int count) {
        if (start + count > s.Length) return false;
        for (int i = start; i < start + count; i++) {
            if (!IsDigit(s[i])) return false;
        }
        return true;
    }

    private static int Number(ReadOnlySpan<char> s, int start, int count) {
        int value = 0;
        for (int i = start; i < start + count; i++) value = value * 10 + (s[i] - '0');
        return value;
    }

    // ---- Strings --------------------------------------------------------------------------------------------------

    /// <summary>Returns the style of a quoted token from its raw text.</summary>
    public static TomlStringStyle StyleOf(ReadOnlySpan<char> raw) {
        if (raw.Length == 0) return TomlStringStyle.Bare;
        bool triple = raw.Length >= 6 && raw[1] == raw[0] && raw[2] == raw[0];
        return raw[0] switch {
            '"' => triple ? TomlStringStyle.MultilineBasic : TomlStringStyle.Basic,
            '\'' => triple ? TomlStringStyle.MultilineLiteral : TomlStringStyle.Literal,
            _ => TomlStringStyle.Bare
        };
    }

    /// <summary>
    /// Decodes the raw text of a validated key or string (with its quotes) to its value. Short values without escapes
    /// come from <see cref="TomlStringCache"/>, since keys and many values repeat.
    /// </summary>
    public static string Decode(ReadOnlySpan<char> raw, TomlStringStyle style) {
        switch (style) {
            case TomlStringStyle.Bare:
                return TomlStringCache.Get(raw);
            case TomlStringStyle.Literal:
                return TomlStringCache.Get(raw.Slice(1, raw.Length - 2));
            case TomlStringStyle.Basic:
                return DecodeBasic(raw.Slice(1, raw.Length - 2), multiline: false);
            case TomlStringStyle.MultilineLiteral: {
                ReadOnlySpan<char> body = TrimFirstNewline(raw.Slice(3, raw.Length - 6));
                return body.IndexOf('\r') < 0 ? body.ToString() : NormalizeNewlines(body);
            }
            default:
                return DecodeBasic(TrimFirstNewline(raw.Slice(3, raw.Length - 6)), multiline: true);
        }
    }

    private static ReadOnlySpan<char> TrimFirstNewline(ReadOnlySpan<char> body) {
        if (body.Length > 0 && body[0] == '\n') return body.Slice(1);
        if (body.Length > 1 && body[0] == '\r' && body[1] == '\n') return body.Slice(2);
        return body;
    }

    private static string NormalizeNewlines(ReadOnlySpan<char> body) {
        var text = new StringBuilder(body.Length);
        for (int i = 0; i < body.Length; i++) {
            if (body[i] == '\r' && i + 1 < body.Length && body[i + 1] == '\n') continue;
            text.Append(body[i]);
        }
        return text.ToString();
    }

    private static string DecodeBasic(ReadOnlySpan<char> body, bool multiline) {
        int first = multiline ? body.IndexOfAny('\\', '\r') : body.IndexOf('\\');
        if (first < 0) return TomlStringCache.Get(body);
        var text = new StringBuilder(body.Length);
        text.Append(body.Slice(0, first).ToString());
        for (int i = first; i < body.Length; i++) {
            char ch = body[i];
            if (ch == '\r' && multiline && i + 1 < body.Length && body[i + 1] == '\n') continue;
            if (ch != '\\') {
                text.Append(ch);
                continue;
            }
            char escape = body[++i];
            switch (escape) {
                case 'b': text.Append('\b'); break;
                case 't': text.Append('\t'); break;
                case 'n': text.Append('\n'); break;
                case 'f': text.Append('\f'); break;
                case 'r': text.Append('\r'); break;
                case 'e': text.Append('\u001B'); break;
                case '"': text.Append('"'); break;
                case '\\': text.Append('\\'); break;
                case 'x': AppendScalar(text, Hex(body, i + 1, 2)); i += 2; break;
                case 'u': AppendScalar(text, Hex(body, i + 1, 4)); i += 4; break;
                case 'U': AppendScalar(text, Hex(body, i + 1, 8)); i += 8; break;
                default:
                    // A line-ending backslash: skip the whitespace and newlines that follow.
                    i++;
                    while (i < body.Length && (body[i] == ' ' || body[i] == '\t' || body[i] == '\n' || body[i] == '\r')) i++;
                    i--;
                    break;
            }
        }
        return text.ToString();
    }

    private static int Hex(ReadOnlySpan<char> s, int start, int count) {
        int value = 0;
        for (int i = start; i < start + count; i++) value = value * 16 + HexValue(s[i]);
        return value;
    }

    private static void AppendScalar(StringBuilder text, int scalar) {
        if (scalar < 0x10000) {
            text.Append((char)scalar);
        } else {
            scalar -= 0x10000;
            text.Append((char)(0xD800 + (scalar >> 10)));
            text.Append((char)(0xDC00 + (scalar & 0x3FF)));
        }
    }

    /// <summary>
    /// Validates an escape at <paramref name="i"/> (the backslash) of a basic string. Returns the escape length, or 0 when
    /// it is invalid. Line-ending backslashes are handled by the scanner.
    /// </summary>
    public static int EscapeLength(ReadOnlySpan<char> s, int i) {
        if (i + 1 >= s.Length) return 0;
        switch (s[i + 1]) {
            case 'b': case 't': case 'n': case 'f': case 'r': case 'e': case '"': case '\\':
                return 2;
            case 'x': return IsScalarEscape(s, i + 2, 2) ? 4 : 0;
            case 'u': return IsScalarEscape(s, i + 2, 4) ? 6 : 0;
            case 'U': return IsScalarEscape(s, i + 2, 8) ? 10 : 0;
            default: return 0;
        }
    }

    private static bool IsScalarEscape(ReadOnlySpan<char> s, int start, int count) {
        if (start + count > s.Length) return false;
        long value = 0;
        for (int i = start; i < start + count; i++) {
            int digit = HexValue(s[i]);
            if (digit < 0) return false;
            value = value * 16 + digit;
        }
        return value <= 0x10FFFF && (value < 0xD800 || value > 0xDFFF);
    }
}
