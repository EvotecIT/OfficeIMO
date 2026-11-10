namespace OfficeIMO.Yaml;

/// <summary>The type an untagged plain scalar resolves to under a schema.</summary>
internal enum YamlScalarKind : byte {
    String = 0,
    Null,
    Boolean,
    Integer,
    Float,
    Timestamp,
    Merge,
    Value
}

/// <summary>
/// Implements the scalar resolution tables of the YAML 1.2.2 failsafe, JSON and core schemas (section 10) and the
/// YAML 1.1 types (https://yaml.org/type/). Matching works on spans and never allocates. The document model reuses
/// it for typed reads, so resolution stays identical between the reader and the model.
/// </summary>
internal static class YamlSchemaResolver {
    public const string TagPrefix = "tag:yaml.org,2002:";

    /// <summary>Resolves an untagged plain scalar.</summary>
    public static YamlScalarKind Resolve(ReadOnlySpan<char> s, YamlSchema schema) {
        switch (schema) {
            case YamlSchema.Failsafe:
                return YamlScalarKind.String;
            case YamlSchema.Json:
                if (Is(s, "null")) return YamlScalarKind.Null;
                if (Is(s, "true") || Is(s, "false")) return YamlScalarKind.Boolean;
                if (IsJsonInteger(s)) return YamlScalarKind.Integer;
                if (IsJsonFloat(s)) return YamlScalarKind.Float;
                return YamlScalarKind.String;
            case YamlSchema.Yaml11:
                if (IsNull11(s)) return YamlScalarKind.Null;
                if (TryBoolean11(s, out _)) return YamlScalarKind.Boolean;
                if (IsInteger11(s)) return YamlScalarKind.Integer;
                if (IsFloat11(s)) return YamlScalarKind.Float;
                if (IsTimestamp11(s)) return YamlScalarKind.Timestamp;
                if (Is(s, "<<")) return YamlScalarKind.Merge;
                if (Is(s, "=")) return YamlScalarKind.Value;
                return YamlScalarKind.String;
            default:
                if (IsNullCore(s)) return YamlScalarKind.Null;
                if (TryBooleanCore(s, out _)) return YamlScalarKind.Boolean;
                if (IsIntegerCore(s)) return YamlScalarKind.Integer;
                if (IsFloatCore(s)) return YamlScalarKind.Float;
                return YamlScalarKind.String;
        }
    }

    /// <summary>Maps a resolved tag (<c>tag:yaml.org,2002:int</c> and so on) to a kind; other tags are not resolvable.</summary>
    public static bool TryGetTagKind(string tag, out YamlScalarKind kind) {
        kind = YamlScalarKind.String;
        if (!tag.StartsWith(TagPrefix, StringComparison.Ordinal)) return false;
        ReadOnlySpan<char> name = tag.AsSpan(TagPrefix.Length);
        if (Is(name, "str")) { kind = YamlScalarKind.String; return true; }
        if (Is(name, "null")) { kind = YamlScalarKind.Null; return true; }
        if (Is(name, "bool")) { kind = YamlScalarKind.Boolean; return true; }
        if (Is(name, "int")) { kind = YamlScalarKind.Integer; return true; }
        if (Is(name, "float")) { kind = YamlScalarKind.Float; return true; }
        if (Is(name, "timestamp")) { kind = YamlScalarKind.Timestamp; return true; }
        return false;
    }

    // ------------------------------------------------------------ null

    public static bool IsNull(ReadOnlySpan<char> s, YamlSchema schema) => schema switch {
        YamlSchema.Failsafe => false,
        YamlSchema.Json => Is(s, "null"),
        YamlSchema.Yaml11 => IsNull11(s),
        _ => IsNullCore(s)
    };

    private static bool IsNullCore(ReadOnlySpan<char> s) =>
        s.Length == 0 || Is(s, "~") || Is(s, "null") || Is(s, "Null") || Is(s, "NULL");

    private static bool IsNull11(ReadOnlySpan<char> s) => IsNullCore(s);

    // ------------------------------------------------------------ boolean

    public static bool TryParseBoolean(ReadOnlySpan<char> s, YamlSchema schema, out bool value) {
        value = false;
        switch (schema) {
            case YamlSchema.Failsafe:
                return false;
            case YamlSchema.Json:
                if (Is(s, "true")) { value = true; return true; }
                return Is(s, "false");
            case YamlSchema.Yaml11:
                return TryBoolean11(s, out value);
            default:
                return TryBooleanCore(s, out value);
        }
    }

    private static bool TryBooleanCore(ReadOnlySpan<char> s, out bool value) {
        value = false;
        if (Is(s, "true") || Is(s, "True") || Is(s, "TRUE")) { value = true; return true; }
        return Is(s, "false") || Is(s, "False") || Is(s, "FALSE");
    }

    private static bool TryBoolean11(ReadOnlySpan<char> s, out bool value) {
        value = true;
        if (Is(s, "y") || Is(s, "Y") || Is(s, "yes") || Is(s, "Yes") || Is(s, "YES")
            || Is(s, "true") || Is(s, "True") || Is(s, "TRUE")
            || Is(s, "on") || Is(s, "On") || Is(s, "ON")) return true;
        value = false;
        return Is(s, "n") || Is(s, "N") || Is(s, "no") || Is(s, "No") || Is(s, "NO")
            || Is(s, "false") || Is(s, "False") || Is(s, "FALSE")
            || Is(s, "off") || Is(s, "Off") || Is(s, "OFF");
    }

    // ------------------------------------------------------------ integers

    /// <summary><c>-?(0|[1-9][0-9]*)</c></summary>
    private static bool IsJsonInteger(ReadOnlySpan<char> s) {
        int i = 0;
        if (i < s.Length && s[i] == '-') i++;
        if (i >= s.Length) return false;
        if (s[i] == '0') return i + 1 == s.Length;
        return AllDigits(s.Slice(i)) && s[i] != '0';
    }

    /// <summary><c>[-+]?[0-9]+</c> | <c>0o[0-7]+</c> | <c>0x[0-9a-fA-F]+</c></summary>
    private static bool IsIntegerCore(ReadOnlySpan<char> s) {
        if (s.Length > 2 && s[0] == '0' && s[1] == 'o') return AllInBase(s.Slice(2), 8, underscores: false);
        if (s.Length > 2 && s[0] == '0' && s[1] == 'x') return AllInBase(s.Slice(2), 16, underscores: false);
        int i = s.Length > 0 && (s[0] == '-' || s[0] == '+') ? 1 : 0;
        return i < s.Length && AllDigits(s.Slice(i));
    }

    /// <summary>
    /// <c>[-+]?0b[0-1_]+</c> | <c>[-+]?0[0-7_]+</c> | <c>[-+]?(0|[1-9][0-9_]*)</c> | <c>[-+]?0x[0-9a-fA-F_]+</c> |
    /// <c>[-+]?[1-9][0-9_]*(:[0-5]?[0-9])+</c>
    /// </summary>
    private static bool IsInteger11(ReadOnlySpan<char> s) {
        int i = s.Length > 0 && (s[0] == '-' || s[0] == '+') ? 1 : 0;
        ReadOnlySpan<char> body = s.Slice(i);
        if (body.Length == 0) return false;
        if (body.Length > 2 && body[0] == '0' && body[1] == 'b') return AllInBase(body.Slice(2), 2, underscores: true);
        if (body.Length > 2 && body[0] == '0' && body[1] == 'x') return AllInBase(body.Slice(2), 16, underscores: true);
        if (body[0] == '0') return body.Length == 1 || AllInBase(body.Slice(1), 8, underscores: true);
        if (body[0] < '1' || body[0] > '9') return false;
        int colon = body.IndexOf(':');
        if (colon < 0) return AllDigitsOrUnderscore(body);
        return AllDigitsOrUnderscore(body.Slice(0, colon)) && IsSexagesimalTail(body.Slice(colon), allowFraction: false);
    }

    /// <summary>Matches <c>(:[0-5]?[0-9])+</c>, optionally followed by <c>\.[0-9_]*</c>.</summary>
    private static bool IsSexagesimalTail(ReadOnlySpan<char> s, bool allowFraction) {
        int i = 0;
        int groups = 0;
        while (i < s.Length && s[i] == ':') {
            i++;
            int digits = 0;
            int start = i;
            while (i < s.Length && IsDigit(s[i]) && digits < 2) { i++; digits++; }
            if (digits == 0) return false;
            if (digits == 2 && s[start] > '5') return false;
            groups++;
        }
        if (groups == 0) return false;
        if (i == s.Length) return true;
        if (!allowFraction || s[i] != '.') return false;
        return AllDigitsOrUnderscore(s.Slice(i + 1), allowEmpty: true);
    }

    public static bool TryParseInt64(ReadOnlySpan<char> s, YamlSchema schema, out long value) {
        value = 0;
        switch (schema) {
            case YamlSchema.Failsafe:
                return false;
            case YamlSchema.Json:
                return IsJsonInteger(s) && TryParseDecimalDigits(s, out value);
            case YamlSchema.Yaml11:
                return IsInteger11(s) && TryParseInteger11(s, out value);
            default:
                if (!IsIntegerCore(s)) return false;
                if (s.Length > 2 && s[0] == '0' && s[1] == 'o') return TryParseBase(s.Slice(2), 8, false, out value);
                if (s.Length > 2 && s[0] == '0' && s[1] == 'x') return TryParseBase(s.Slice(2), 16, false, out value);
                return TryParseDecimalDigits(s, out value);
        }
    }

    private static bool TryParseInteger11(ReadOnlySpan<char> s, out long value) {
        value = 0;
        bool negative = s[0] == '-';
        ReadOnlySpan<char> body = s[0] == '-' || s[0] == '+' ? s.Slice(1) : s;
        ulong magnitude;
        if (body.Length > 2 && body[0] == '0' && body[1] == 'b') {
            if (!TryParseUnsigned(body.Slice(2), 2, out magnitude)) return false;
        } else if (body.Length > 2 && body[0] == '0' && body[1] == 'x') {
            if (!TryParseUnsigned(body.Slice(2), 16, out magnitude)) return false;
        } else if (body.Length > 1 && body[0] == '0') {
            if (!TryParseUnsigned(body.Slice(1), 8, out magnitude)) return false;
        } else if (body.IndexOf(':') >= 0) {
            magnitude = 0;
            foreach (ReadOnlySpan<char> part in new SexagesimalParts(body)) {
                if (!TryParseUnsigned(part, 10, out ulong group)) return false;
                if (magnitude > (ulong.MaxValue - group) / 60) return false;
                magnitude = magnitude * 60 + group;
            }
        } else if (!TryParseUnsigned(body, 10, out magnitude)) {
            return false;
        }
        return ApplySign(magnitude, negative, out value);
    }

    private static bool TryParseDecimalDigits(ReadOnlySpan<char> s, out long value) {
        value = 0;
        bool negative = s[0] == '-';
        ReadOnlySpan<char> body = s[0] == '-' || s[0] == '+' ? s.Slice(1) : s;
        return TryParseUnsigned(body, 10, out ulong magnitude) && ApplySign(magnitude, negative, out value);
    }

    private static bool TryParseBase(ReadOnlySpan<char> s, int numberBase, bool negative, out long value) {
        value = 0;
        return TryParseUnsigned(s, numberBase, out ulong magnitude) && ApplySign(magnitude, negative, out value);
    }

    private static bool ApplySign(ulong magnitude, bool negative, out long value) {
        value = 0;
        if (negative) {
            if (magnitude > (ulong)long.MaxValue + 1) return false;
            value = magnitude == (ulong)long.MaxValue + 1 ? long.MinValue : -(long)magnitude;
            return true;
        }
        if (magnitude > long.MaxValue) return false;
        value = (long)magnitude;
        return true;
    }

    /// <summary>Parses digits in the base, skipping underscores (only present in YAML 1.1 forms).</summary>
    private static bool TryParseUnsigned(ReadOnlySpan<char> s, int numberBase, out ulong value) {
        value = 0;
        bool any = false;
        foreach (char ch in s) {
            if (ch == '_') continue;
            int digit = YamlChars.HexValue(ch);
            if (digit < 0 || digit >= numberBase) return false;
            if (value > (ulong.MaxValue - (ulong)digit) / (ulong)numberBase) return false;
            value = value * (ulong)numberBase + (ulong)digit;
            any = true;
        }
        return any;
    }

    // ------------------------------------------------------------ floats

    /// <summary><c>-?(0|[1-9][0-9]*)(\.[0-9]*)?([eE][-+]?[0-9]+)?</c></summary>
    private static bool IsJsonFloat(ReadOnlySpan<char> s) {
        int i = 0;
        if (i < s.Length && s[i] == '-') i++;
        if (i >= s.Length) return false;
        if (s[i] == '0') i++;
        else if (s[i] >= '1' && s[i] <= '9') { while (i < s.Length && IsDigit(s[i])) i++; } else return false;
        if (i < s.Length && s[i] == '.') { i++; while (i < s.Length && IsDigit(s[i])) i++; }
        return MatchExponent(s, i, signRequired: false);
    }

    /// <summary><c>[-+]?(\.[0-9]+|[0-9]+(\.[0-9]*)?)([eE][-+]?[0-9]+)?</c> | <c>[-+]?\.(inf|Inf|INF)</c> | <c>\.nan|\.NaN|\.NAN</c></summary>
    private static bool IsFloatCore(ReadOnlySpan<char> s) {
        if (IsNaN(s)) return true;
        int i = s.Length > 0 && (s[0] == '-' || s[0] == '+') ? 1 : 0;
        if (IsInfinity(s.Slice(i))) return true;
        int digits = 0;
        while (i < s.Length && IsDigit(s[i])) { i++; digits++; }
        if (i < s.Length && s[i] == '.') {
            i++;
            int fraction = 0;
            while (i < s.Length && IsDigit(s[i])) { i++; fraction++; }
            if (digits == 0 && fraction == 0) return false;
        } else if (digits == 0) {
            return false;
        }
        return MatchExponent(s, i, signRequired: false);
    }

    /// <summary>
    /// <c>[-+]?([0-9][0-9_]*)?\.[0-9.]*([eE][-+][0-9]+)?</c> | <c>[-+]?[0-9][0-9_]*(:[0-5]?[0-9])+\.[0-9_]*</c> |
    /// <c>[-+]?\.(inf|Inf|INF)</c> | <c>\.(nan|NaN|NAN)</c>. A match without any digit (such as <c>.</c>) is not a number.
    /// </summary>
    private static bool IsFloat11(ReadOnlySpan<char> s) {
        if (IsNaN(s)) return true;
        int i = s.Length > 0 && (s[0] == '-' || s[0] == '+') ? 1 : 0;
        if (IsInfinity(s.Slice(i))) return true;
        int start = i;
        int colon = s.IndexOf(':');
        if (colon >= 0) {
            if (i >= s.Length || !IsDigit(s[i])) return false;
            if (!AllDigitsOrUnderscore(s.Slice(i, colon - i))) return false;
            int dot = s.Slice(colon).IndexOf('.');
            if (dot < 0) return false;
            return IsSexagesimalTail(s.Slice(colon), allowFraction: true);
        }
        bool anyDigit = false;
        if (i < s.Length && IsDigit(s[i])) {
            while (i < s.Length && (IsDigit(s[i]) || s[i] == '_')) { anyDigit |= s[i] != '_'; i++; }
        }
        if (i >= s.Length || s[i] != '.') return false;
        i++;
        while (i < s.Length && (IsDigit(s[i]) || s[i] == '.')) { anyDigit |= s[i] != '.'; i++; }
        if (!anyDigit || i == start) return false;
        return MatchExponent(s, i, signRequired: true);
    }

    private static bool MatchExponent(ReadOnlySpan<char> s, int i, bool signRequired) {
        if (i == s.Length) return true;
        if (s[i] != 'e' && s[i] != 'E') return false;
        i++;
        if (i < s.Length && (s[i] == '-' || s[i] == '+')) i++;
        else if (signRequired) return false;
        int digits = 0;
        while (i < s.Length && IsDigit(s[i])) { i++; digits++; }
        return digits > 0 && i == s.Length;
    }

    private static bool IsNaN(ReadOnlySpan<char> s) => Is(s, ".nan") || Is(s, ".NaN") || Is(s, ".NAN");

    private static bool IsInfinity(ReadOnlySpan<char> s) => Is(s, ".inf") || Is(s, ".Inf") || Is(s, ".INF");

    public static bool TryParseDouble(ReadOnlySpan<char> s, YamlSchema schema, out double value) {
        value = 0;
        switch (schema) {
            case YamlSchema.Failsafe:
                return false;
            case YamlSchema.Json:
                if (!IsJsonFloat(s)) return false;
                return ParseFloat(s, out value);
            case YamlSchema.Yaml11:
                if (IsInteger11(s)) {
                    if (TryParseInteger11(s, out long whole)) { value = whole; return true; }
                    return ParseFloat(s, out value);
                }
                if (!IsFloat11(s)) return false;
                if (s.IndexOf(':') >= 0) return TryParseSexagesimalFloat(s, out value);
                return ParseSpecialOrFloat(s, out value);
            default:
                if (IsIntegerCore(s)) {
                    if (TryParseInt64(s, YamlSchema.Core, out long whole)) { value = whole; return true; }
                    if (s.Length > 2 && s[0] == '0' && (s[1] == 'o' || s[1] == 'x')) return TryParseBigBase(s.Slice(2), s[1] == 'o' ? 8 : 16, out value);
                    return ParseFloat(s, out value);
                }
                if (!IsFloatCore(s)) return false;
                return ParseSpecialOrFloat(s, out value);
        }
    }

    public static bool TryParseDecimal(ReadOnlySpan<char> s, YamlSchema schema, out decimal value) {
        value = 0;
        if (schema == YamlSchema.Failsafe) return false;
        bool isInteger = schema switch {
            YamlSchema.Json => IsJsonInteger(s),
            YamlSchema.Yaml11 => IsInteger11(s),
            _ => IsIntegerCore(s)
        };
        if (isInteger) {
            if (TryParseInt64(s, schema, out long whole)) { value = whole; return true; }
            bool decimalDigits = schema == YamlSchema.Json
                || (schema == YamlSchema.Core && !(s.Length > 2 && s[0] == '0' && (s[1] == 'o' || s[1] == 'x')))
                || (schema == YamlSchema.Yaml11 && Unprefixed11(s));
            return decimalDigits && ParseDecimalText(s, out value);
        }
        bool isFloat = schema switch {
            YamlSchema.Json => IsJsonFloat(s),
            YamlSchema.Yaml11 => IsFloat11(s),
            _ => IsFloatCore(s)
        };
        if (!isFloat) return false;
        int i = s.Length > 0 && (s[0] == '-' || s[0] == '+') ? 1 : 0;
        if (IsNaN(s) || IsInfinity(s.Slice(i))) return false;
        if (schema == YamlSchema.Yaml11 && s.IndexOf(':') >= 0) {
            if (!TryParseSexagesimalFloat(s, out double approximate)) return false;
            value = (decimal)approximate;
            return true;
        }
        return ParseDecimalText(s, out value);
    }

    private static bool Unprefixed11(ReadOnlySpan<char> s) {
        ReadOnlySpan<char> body = s.Length > 0 && (s[0] == '-' || s[0] == '+') ? s.Slice(1) : s;
        return body.Length > 0 && body[0] != '0' && body.IndexOf(':') < 0;
    }

    private static bool ParseSpecialOrFloat(ReadOnlySpan<char> s, out double value) {
        if (IsNaN(s)) { value = double.NaN; return true; }
        int i = s.Length > 0 && (s[0] == '-' || s[0] == '+') ? 1 : 0;
        if (IsInfinity(s.Slice(i))) {
            value = s[0] == '-' ? double.NegativeInfinity : double.PositiveInfinity;
            return true;
        }
        return ParseFloat(s, out value);
    }

    private static bool TryParseSexagesimalFloat(ReadOnlySpan<char> s, out double value) {
        value = 0;
        bool negative = s[0] == '-';
        ReadOnlySpan<char> body = s[0] == '-' || s[0] == '+' ? s.Slice(1) : s;
        int dot = body.LastIndexOf('.');
        ReadOnlySpan<char> whole = dot >= 0 ? body.Slice(0, dot) : body;
        double result = 0;
        foreach (ReadOnlySpan<char> part in new SexagesimalParts(whole)) {
            if (!TryParseUnsigned(part, 10, out ulong group)) return false;
            result = result * 60 + group;
        }
        if (dot >= 0 && dot + 1 < body.Length) {
            if (!ParseFloat(("0." + body.Slice(dot + 1).ToString()).AsSpan(), out double fraction)) return false;
            result += fraction;
        }
        value = negative ? -result : result;
        return true;
    }

    private static bool TryParseBigBase(ReadOnlySpan<char> s, int numberBase, out double value) {
        value = 0;
        foreach (char ch in s) {
            int digit = YamlChars.HexValue(ch);
            if (digit < 0 || digit >= numberBase) return false;
            value = value * numberBase + digit;
        }
        return true;
    }

    private static bool ParseFloat(ReadOnlySpan<char> s, out double value) {
        Span<char> buffer = s.Length <= 128 ? stackalloc char[s.Length] : new char[s.Length];
        int n = CopyWithoutUnderscores(s, buffer);
#if NET8_0_OR_GREATER
        return double.TryParse(buffer.Slice(0, n), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
#else
        return double.TryParse(buffer.Slice(0, n).ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
#endif
    }

    private static bool ParseDecimalText(ReadOnlySpan<char> s, out decimal value) {
        Span<char> buffer = s.Length <= 128 ? stackalloc char[s.Length] : new char[s.Length];
        int n = CopyWithoutUnderscores(s, buffer);
#if NET8_0_OR_GREATER
        return decimal.TryParse(buffer.Slice(0, n), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
#else
        return decimal.TryParse(buffer.Slice(0, n).ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
#endif
    }

    private static int CopyWithoutUnderscores(ReadOnlySpan<char> s, Span<char> buffer) {
        int n = 0;
        foreach (char ch in s) {
            if (ch != '_') buffer[n++] = ch;
        }
        return n;
    }

    // ------------------------------------------------------------ timestamps (YAML 1.1)

    /// <summary>
    /// <c>[0-9]{4}-[0-9]{2}-[0-9]{2}</c> |
    /// <c>[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}([Tt]|[ \t]+)[0-9]{1,2}:[0-9]{2}:[0-9]{2}(\.[0-9]*)?(([ \t]*)Z|[-+][0-9]{1,2}(:[0-9]{2})?)?</c>
    /// </summary>
    private static bool IsTimestamp11(ReadOnlySpan<char> s) {
        int i = 0;
        if (!Digits(s, ref i, 4, 4) || !Char(s, ref i, '-')) return false;
        int monthStart = i;
        if (!Digits(s, ref i, 1, 2) || !Char(s, ref i, '-')) return false;
        int monthDigits = i - monthStart - 1;
        int dayStart = i;
        if (!Digits(s, ref i, 1, 2)) return false;
        int dayDigits = i - dayStart;
        if (i == s.Length) return monthDigits == 2 && dayDigits == 2;
        if (s[i] == 'T' || s[i] == 't') i++;
        else {
            int blanks = 0;
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) { i++; blanks++; }
            if (blanks == 0) return false;
        }
        if (!Digits(s, ref i, 1, 2) || !Char(s, ref i, ':') || !Digits(s, ref i, 2, 2) || !Char(s, ref i, ':') || !Digits(s, ref i, 2, 2)) return false;
        if (i < s.Length && s[i] == '.') { i++; while (i < s.Length && IsDigit(s[i])) i++; }
        if (i == s.Length) return true;
        int save = i;
        while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) i++;
        if (i < s.Length && s[i] == 'Z') return i + 1 == s.Length;
        i = save;
        if (s[i] != '-' && s[i] != '+') return false;
        i++;
        if (!Digits(s, ref i, 1, 2)) return false;
        if (i == s.Length) return true;
        return Char(s, ref i, ':') && Digits(s, ref i, 2, 2) && i == s.Length;
    }

    private static bool Digits(ReadOnlySpan<char> s, ref int i, int min, int max) {
        int count = 0;
        while (i < s.Length && count < max && IsDigit(s[i])) { i++; count++; }
        return count >= min;
    }

    private static bool Char(ReadOnlySpan<char> s, ref int i, char ch) {
        if (i < s.Length && s[i] == ch) { i++; return true; }
        return false;
    }

    // ------------------------------------------------------------ helpers

    private static bool Is(ReadOnlySpan<char> s, string literal) => s.SequenceEqual(literal.AsSpan());

    private static bool IsDigit(char ch) => ch >= '0' && ch <= '9';

    private static bool AllDigits(ReadOnlySpan<char> s) {
        if (s.Length == 0) return false;
        foreach (char ch in s) if (!IsDigit(ch)) return false;
        return true;
    }

    private static bool AllDigitsOrUnderscore(ReadOnlySpan<char> s, bool allowEmpty = false) {
        if (s.Length == 0) return allowEmpty;
        foreach (char ch in s) if (!IsDigit(ch) && ch != '_') return false;
        return true;
    }

    private static bool AllInBase(ReadOnlySpan<char> s, int numberBase, bool underscores) {
        if (s.Length == 0) return false;
        foreach (char ch in s) {
            if (underscores && ch == '_') continue;
            int digit = YamlChars.HexValue(ch);
            if (digit < 0 || digit >= numberBase) return false;
        }
        return true;
    }

    /// <summary>Enumerates the colon-separated groups of a sexagesimal number without allocating.</summary>
    private ref struct SexagesimalParts {
        private ReadOnlySpan<char> _rest;
        private bool _done;

        public SexagesimalParts(ReadOnlySpan<char> s) {
            _rest = s;
            _done = false;
            Current = default;
        }

        public ReadOnlySpan<char> Current { get; private set; }

        public readonly SexagesimalParts GetEnumerator() => this;

        public bool MoveNext() {
            if (_done) return false;
            int colon = _rest.IndexOf(':');
            if (colon < 0) {
                Current = _rest;
                _done = true;
                return true;
            }
            Current = _rest.Slice(0, colon);
            _rest = _rest.Slice(colon + 1);
            return true;
        }
    }
}
