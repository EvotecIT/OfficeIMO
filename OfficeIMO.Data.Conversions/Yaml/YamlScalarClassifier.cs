using System.Numerics;
using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Resolves YAML scalars to data kinds through the engine's public typed accessors, so conversion follows exactly the
/// schema the engine applies: the 1.2 core schema by default, YAML 1.1 rules (including timestamps) only under the 1.1
/// schema. Numbers keep their spelling when it is valid JSON (<c>1.10</c>); other spellings (<c>0x1F</c>, <c>.5</c>)
/// become canonical text.
/// </summary>
internal static class YamlScalarClassifier {
    internal const string TagPrefix = "tag:yaml.org,2002:";

    /// <summary>A view of one scalar from either the pull reader or the document model.</summary>
    internal interface IScalar {
        string? Tag { get; }
        YamlStyle Style { get; }
        ReadOnlySpan<char> Text { get; }
        bool IsNull();
        bool TryGetBoolean(out bool value);
        bool TryGetInt64(out long value);
        bool TryGetDouble(out double value);
    }

    /// <summary>Classifies a value scalar and leaves its canonical text in <paramref name="text"/>.</summary>
    public static DataScalarKind Classify<TScalar>(ref TScalar scalar, bool yaml11, DataConversionContext context, out ReadOnlySpan<char> text)
        where TScalar : struct, IScalar {
        string? tag = scalar.Tag;
        if (tag != null) return ClassifyTagged(ref scalar, tag, context, out text);
        if (scalar.Style != YamlStyle.Plain) {
            text = scalar.Text;
            return DataScalarKind.String;
        }

        ReadOnlySpan<char> raw = scalar.Text;
        if (raw.Length > 0 && !MayResolve(raw[0])) {
            text = raw;
            return DataScalarKind.String;
        }

        if (scalar.IsNull()) {
            text = default;
            return DataScalarKind.Null;
        }

        if (scalar.TryGetBoolean(out bool boolean)) {
            text = boolean ? "true".AsSpan() : "false".AsSpan();
            return DataScalarKind.Boolean;
        }

        if (scalar.TryGetInt64(out long integer)) {
            text = IntegerText(integer, raw);
            return DataScalarKind.Integer;
        }

        if (scalar.TryGetDouble(out double number)) return Number(raw, number, out text);
        if (yaml11 && TryTimestamp(raw, out DataScalarKind dateKind, out string? normalized)) {
            text = normalized.AsSpan();
            return dateKind;
        }

        text = raw;
        return DataScalarKind.String;
    }

    /// <summary>Resolves a mapping key to the text a JSON-like target uses, reporting keys that were not strings.</summary>
    public static ReadOnlySpan<char> KeyText<TScalar>(ref TScalar scalar, bool yaml11, DataConversionContext context) where TScalar : struct, IScalar {
        DataScalarKind kind = Classify(ref scalar, yaml11, context, out ReadOnlySpan<char> text);
        if (kind == DataScalarKind.String) return text;
        context.Report(OfficeDiagnosticCodes.KeyTypeChanged, OfficeConversionLossKind.Approximation,
            "Mapping keys that were numbers, booleans, nulls or dates became string keys.");
        return kind == DataScalarKind.Null ? "null".AsSpan() : text;
    }

    private static DataScalarKind ClassifyTagged<TScalar>(ref TScalar scalar, string tag, DataConversionContext context, out ReadOnlySpan<char> text)
        where TScalar : struct, IScalar {
        if (tag.StartsWith(TagPrefix, StringComparison.Ordinal)) {
            switch (tag.Substring(TagPrefix.Length)) {
                case "str":
                case "binary":
                    text = scalar.Text;
                    return DataScalarKind.String;
                case "null":
                    text = default;
                    return DataScalarKind.Null;
                case "bool":
                    if (scalar.TryGetBoolean(out bool boolean)) {
                        text = boolean ? "true".AsSpan() : "false".AsSpan();
                        return DataScalarKind.Boolean;
                    }

                    break;
                case "int":
                    if (scalar.TryGetInt64(out long integer)) {
                        text = IntegerText(integer, scalar.Text);
                        return DataScalarKind.Integer;
                    }

                    if (scalar.TryGetDouble(out double big)) return Number(scalar.Text, big, out text);
                    break;
                case "float":
                    if (scalar.TryGetDouble(out double number)) {
                        DataScalarKind kind = Number(scalar.Text, number, out text);
                        if (kind == DataScalarKind.Integer) return Float(number, out text);
                        return kind;
                    }

                    break;
                case "timestamp":
                    if (TryTimestamp(scalar.Text, out DataScalarKind dateKind, out string? normalized)) {
                        text = normalized.AsSpan();
                        return dateKind;
                    }

                    break;
            }

            throw new InvalidDataException("The scalar '" + scalar.Text.ToString() + "' does not match its tag " + tag + ".");
        }

        if (tag != "!") {
            context.Report(OfficeDiagnosticCodes.YamlTagDropped, OfficeConversionLossKind.Approximation,
                "Custom YAML tags (such as " + tag + ") were dropped; their values were kept as text.");
        }

        text = scalar.Text;
        return DataScalarKind.String;
    }

    /// <summary>Reports a custom tag on a mapping or sequence.</summary>
    public static void CheckCollectionTag(string? tag, DataConversionContext context) {
        if (tag == null || tag == "!" || tag.StartsWith(TagPrefix, StringComparison.Ordinal)) return;
        context.Report(OfficeDiagnosticCodes.YamlTagDropped, OfficeConversionLossKind.Approximation,
            "Custom YAML tags (such as " + tag + ") were dropped; their values were kept as text.");
    }

    // Under every schema, a plain scalar that resolves to null, a boolean, a number or a timestamp starts with one of these.
    private static bool MayResolve(char c) =>
        (c >= '0' && c <= '9') || c is '+' or '-' or '.' or '~' or 'n' or 'N' or 't' or 'T' or 'f' or 'F' or 'y' or 'Y' or 'o' or 'O';

    private static ReadOnlySpan<char> IntegerText(long value, ReadOnlySpan<char> raw) {
        if (IsCanonicalInteger(raw)) return raw;
        return value.ToString(CultureInfo.InvariantCulture).AsSpan();
    }

    private static DataScalarKind Number(ReadOnlySpan<char> raw, double value, out ReadOnlySpan<char> text) {
        if (double.IsNaN(value)) {
            text = "NaN".AsSpan();
            return DataScalarKind.Float;
        }

        if (double.IsInfinity(value)) {
            text = value > 0 ? "Infinity".AsSpan() : "-Infinity".AsSpan();
            return DataScalarKind.Float;
        }

        if (TryBigInteger(raw, out string? digits)) {
            text = digits.AsSpan();
            return DataScalarKind.Integer;
        }

        if (JsonNumber.IsValid(raw)) {
            text = raw;
            return DataScalarKind.Float;
        }

        return Float(value, out text);
    }

    private static DataScalarKind Float(double value, out ReadOnlySpan<char> text) {
        string formatted = value.ToString("R", CultureInfo.InvariantCulture);
        if (formatted.IndexOf('.') < 0 && formatted.IndexOf('E') < 0) formatted += ".0";
        text = formatted.AsSpan();
        return DataScalarKind.Float;
    }

    internal static bool IsCanonicalInteger(ReadOnlySpan<char> raw) {
        int i = raw.Length > 0 && raw[0] == '-' ? 1 : 0;
        if (i == raw.Length) return false;
        if (raw[i] == '0' && raw.Length > i + 1) return false;
        for (; i < raw.Length; i++) {
            if (raw[i] < '0' || raw[i] > '9') return false;
        }

        return true;
    }

    // An integer spelling the engine resolved as a number but that does not fit in Int64.
    private static bool TryBigInteger(ReadOnlySpan<char> raw, out string? digits) {
        digits = null;
        int i = 0;
        bool negative = false;
        if (i < raw.Length && (raw[i] == '+' || raw[i] == '-')) negative = raw[i++] == '-';
        if (i >= raw.Length) return false;
        int radix = 10;
        if (raw.Length - i > 2 && raw[i] == '0' && (raw[i + 1] is 'x' or 'o' or 'b')) {
            radix = raw[i + 1] == 'x' ? 16 : raw[i + 1] == 'o' ? 8 : 2;
            i += 2;
        }

        BigInteger value = BigInteger.Zero;
        bool any = false;
        for (; i < raw.Length; i++) {
            char c = raw[i];
            if (c == '_') continue;
            int digit = c >= '0' && c <= '9' ? c - '0' : c >= 'a' && c <= 'f' ? c - 'a' + 10 : c >= 'A' && c <= 'F' ? c - 'A' + 10 : -1;
            if (digit < 0 || digit >= radix) return false;
            value = value * radix + digit;
            any = true;
        }

        if (!any) return false;
        digits = (negative ? -value : value).ToString(CultureInfo.InvariantCulture);
        return true;
    }

    /// <summary>
    /// Recognizes a YAML 1.1 timestamp and normalizes it to RFC 3339. A timestamp without a zone is UTC under YAML 1.1,
    /// so it becomes an offset date-time with <c>Z</c>; a bare date is a local date.
    /// </summary>
    internal static bool TryTimestamp(ReadOnlySpan<char> s, out DataScalarKind kind, out string? normalized) {
        kind = default;
        normalized = null;
        int i = 0;
        if (!Digits(s, ref i, 4, 4, out int year) || !Expect(s, ref i, '-') || !Digits(s, ref i, 1, 2, out int month)
            || !Expect(s, ref i, '-') || !Digits(s, ref i, 1, 2, out int day)) return false;
        if (i == s.Length) {
            if (s.Length != 10) return false;
            kind = DataScalarKind.LocalDate;
            normalized = s.ToString();
            return true;
        }

        if (s[i] == 'T' || s[i] == 't') i++;
        else {
            int spaces = i;
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) i++;
            if (i == spaces) return false;
        }

        if (!Digits(s, ref i, 1, 2, out int hour) || !Expect(s, ref i, ':') || !Digits(s, ref i, 2, 2, out int minute)
            || !Expect(s, ref i, ':') || !Digits(s, ref i, 2, 2, out int second)) return false;
        int fractionStart = i, fractionEnd = i;
        if (i < s.Length && s[i] == '.') {
            i++;
            fractionStart = i;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            fractionEnd = i;
        }

        while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) i++;
        var builder = new StringBuilder(35);
        builder.Append(year.ToString("D4", CultureInfo.InvariantCulture)).Append('-')
            .Append(month.ToString("D2", CultureInfo.InvariantCulture)).Append('-')
            .Append(day.ToString("D2", CultureInfo.InvariantCulture)).Append('T')
            .Append(hour.ToString("D2", CultureInfo.InvariantCulture)).Append(':')
            .Append(minute.ToString("D2", CultureInfo.InvariantCulture)).Append(':')
            .Append(second.ToString("D2", CultureInfo.InvariantCulture));
        if (fractionEnd > fractionStart) builder.Append('.').Append(s.Slice(fractionStart, fractionEnd - fractionStart).ToString());
        if (i == s.Length || (s[i] == 'Z' && i + 1 == s.Length)) {
            builder.Append('Z');
        } else if (s[i] == '+' || s[i] == '-') {
            char sign = s[i++];
            if (!Digits(s, ref i, 1, 2, out int zoneHours)) return false;
            int zoneMinutes = 0;
            if (i < s.Length && s[i] == ':' && (++i > 0) && !Digits(s, ref i, 2, 2, out zoneMinutes)) return false;
            if (i != s.Length) return false;
            builder.Append(sign).Append(zoneHours.ToString("D2", CultureInfo.InvariantCulture)).Append(':')
                .Append(zoneMinutes.ToString("D2", CultureInfo.InvariantCulture));
        } else {
            return false;
        }

        kind = DataScalarKind.OffsetDateTime;
        normalized = builder.ToString();
        return true;
    }

    private static bool Expect(ReadOnlySpan<char> s, ref int i, char c) {
        if (i >= s.Length || s[i] != c) return false;
        i++;
        return true;
    }

    private static bool Digits(ReadOnlySpan<char> s, ref int i, int min, int max, out int value) {
        value = 0;
        int start = i;
        while (i < s.Length && i - start < max && s[i] >= '0' && s[i] <= '9') value = value * 10 + (s[i++] - '0');
        return i - start >= min;
    }
}

/// <summary>Checks the JSON number grammar of RFC 8259.</summary>
internal static class JsonNumber {
    public static bool IsValid(ReadOnlySpan<char> s) {
        int i = 0;
        if (i < s.Length && s[i] == '-') i++;
        if (i >= s.Length) return false;
        if (s[i] == '0') i++;
        else if (s[i] >= '1' && s[i] <= '9') while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
        else return false;
        if (i < s.Length && s[i] == '.') {
            i++;
            int start = i;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            if (i == start) return false;
        }

        if (i < s.Length && (s[i] == 'e' || s[i] == 'E')) {
            i++;
            if (i < s.Length && (s[i] == '+' || s[i] == '-')) i++;
            int start = i;
            while (i < s.Length && s[i] >= '0' && s[i] <= '9') i++;
            if (i == start) return false;
        }

        return i == s.Length;
    }
}
