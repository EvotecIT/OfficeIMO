using System.Globalization;
using System.Numerics;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.Adapters;

/// <summary>
/// YAML 1.2.2 core schema (section 10.3) used by both reference adapters, so JSON equivalence measures each
/// library's structure loading with one shared scalar resolution. Written as scanners rather than regular expressions.
/// </summary>
internal static class YamlCoreSchema {
    private const string Prefix = "tag:yaml.org,2002:";

    /// <summary>Resolves a scalar. <paramref name="tag"/> is null for an untagged node and <c>!</c> for the non-specific tag.</summary>
    public static YamlConformanceNode Resolve(string value, bool isPlain, string? tag) {
        if (string.IsNullOrEmpty(tag)) return isPlain ? ResolvePlain(value) : YamlConformanceNode.String(value);
        switch (tag) {
            case Prefix + "null": return YamlConformanceNode.Null;
            case Prefix + "bool": return TryBool(value, out bool flag) ? YamlConformanceNode.Boolean(flag) : throw new FormatException("Invalid !!bool '" + value + "'.");
            case Prefix + "int": return TryInteger(value, out string? digits) ? YamlConformanceNode.Integer(digits!) : throw new FormatException("Invalid !!int '" + value + "'.");
            case Prefix + "float": return TryFloat(value, out string? number) || TryInteger(value, out number) ? YamlConformanceNode.Float(number!) : throw new FormatException("Invalid !!float '" + value + "'.");
            default: return YamlConformanceNode.String(value);
        }
    }

    private static YamlConformanceNode ResolvePlain(string value) {
        if (value.Length == 0 || value == "~" || value == "null" || value == "Null" || value == "NULL") return YamlConformanceNode.Null;
        if (TryBool(value, out bool flag)) return YamlConformanceNode.Boolean(flag);
        if (TryInteger(value, out string? digits)) return YamlConformanceNode.Integer(digits!);
        if (TryFloat(value, out string? number)) return YamlConformanceNode.Float(number!);
        return YamlConformanceNode.String(value);
    }

    private static bool TryBool(string value, out bool result) {
        result = value == "true" || value == "True" || value == "TRUE";
        return result || value == "false" || value == "False" || value == "FALSE";
    }

    // [-+]?[0-9]+ | 0o[0-7]+ | 0x[0-9a-fA-F]+
    private static bool TryInteger(string value, out string? digits) {
        digits = null;
        if (value.Length > 2 && value[0] == '0' && (value[1] == 'o' || value[1] == 'x')) {
            int radix = value[1] == 'o' ? 8 : 16;
            BigInteger total = BigInteger.Zero;
            for (int i = 2; i < value.Length; i++) {
                int digit = DigitValue(value[i]);
                if (digit < 0 || digit >= radix) return false;
                total = total * radix + digit;
            }

            digits = total.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        int start = value.Length > 0 && (value[0] == '-' || value[0] == '+') ? 1 : 0;
        if (start == value.Length) return false;
        for (int i = start; i < value.Length; i++) {
            if (value[i] < '0' || value[i] > '9') return false;
        }

        digits = BigInteger.Parse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
        return true;
    }

    // [-+]?(\.[0-9]+|[0-9]+(\.[0-9]*)?)([eE][-+]?[0-9]+)? | [-+]?\.(inf|Inf|INF) | \.(nan|NaN|NAN)
    private static bool TryFloat(string value, out string? number) {
        number = null;
        int i = 0;
        bool negative = false;
        if (i < value.Length && (value[i] == '-' || value[i] == '+')) {
            negative = value[i] == '-';
            i++;
        }

        string rest = value.Substring(i);
        if (rest == ".inf" || rest == ".Inf" || rest == ".INF") {
            number = negative ? "-Infinity" : "Infinity";
            return true;
        }

        if (i == 0 && (rest == ".nan" || rest == ".NaN" || rest == ".NAN")) {
            number = "NaN";
            return true;
        }

        int integerDigits = CountDigits(value, ref i);
        int fractionDigits = 0;
        bool hasPoint = false;
        if (i < value.Length && value[i] == '.') {
            hasPoint = true;
            i++;
            fractionDigits = CountDigits(value, ref i);
        }

        if (integerDigits == 0 && fractionDigits == 0) return false;
        if (integerDigits == 0 && !hasPoint) return false;
        if (i < value.Length && (value[i] == 'e' || value[i] == 'E')) {
            i++;
            if (i < value.Length && (value[i] == '-' || value[i] == '+')) i++;
            if (CountDigits(value, ref i) == 0) return false;
        }

        if (i != value.Length) return false;
        // Normalize "1." and ".5" so invariant number parsing accepts them everywhere.
        string text = value.Replace(".e", ".0e").Replace(".E", ".0E");
        if (text.EndsWith(".", StringComparison.Ordinal)) text += "0";
        int point = text.IndexOf('.');
        if (point >= 0 && (point == 0 || !char.IsDigit(text[point - 1]))) text = text.Insert(point, "0");
        number = text;
        return true;
    }

    private static int CountDigits(string value, ref int index) {
        int start = index;
        while (index < value.Length && value[index] >= '0' && value[index] <= '9') index++;
        return index - start;
    }

    private static int DigitValue(char c) =>
        c >= '0' && c <= '9' ? c - '0' : c >= 'a' && c <= 'f' ? c - 'a' + 10 : c >= 'A' && c <= 'F' ? c - 'A' + 10 : -1;
}
