using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

internal static class HtmlCounterStyleFormatter {
    internal const int MaximumGeneratedRepresentationLength = 4096;

    internal static bool TryFormat(int value, string style, out string formatted) =>
        TryFormat(value, style, out formatted, out _);

    internal static bool TryFormat(int value, string style, out string formatted, out bool representationLimited) {
        representationLimited = false;
        string decoded = HtmlCssEscapeDecoder.Decode(style.Trim());
        if (TryUnquote(decoded, out formatted)) return true;
        if (TryFormatSymbolsFunction(value, decoded, out formatted, out representationLimited)) return true;
        if (representationLimited) return true;
        string normalized = decoded.ToLowerInvariant();
        switch (normalized) {
            case "decimal-leading-zero":
                formatted = value >= -9 && value <= 9
                    ? value < 0 ? "-0" + (-value).ToString(CultureInfo.InvariantCulture) : "0" + value.ToString(CultureInfo.InvariantCulture)
                    : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "lower-alpha":
            case "lower-latin":
                formatted = value > 0 ? FormatAlphabetic(value, "abcdefghijklmnopqrstuvwxyz") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "upper-alpha":
            case "upper-latin":
                formatted = value > 0 ? FormatAlphabetic(value, "ABCDEFGHIJKLMNOPQRSTUVWXYZ") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "lower-greek":
                formatted = value > 0 ? FormatAlphabetic(value, "αβγδεζηθικλμνξοπρστυφχψω") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "hiragana":
                formatted = value > 0 ? FormatAlphabetic(value, "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわゐゑをん") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "hiragana-iroha":
                formatted = value > 0 ? FormatAlphabetic(value, "いろはにほへとちりぬるをわかよたれそつねならむうゐのおくやまけふこえてあさきゆめみしゑひもせす") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "katakana":
                formatted = value > 0 ? FormatAlphabetic(value, "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヰヱヲン") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "katakana-iroha":
                formatted = value > 0 ? FormatAlphabetic(value, "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "cjk-heavenly-stem":
                formatted = value > 0 ? FormatAlphabetic(value, "甲乙丙丁戊己庚辛壬癸") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "cjk-earthly-branch":
                formatted = value > 0 ? FormatAlphabetic(value, "子丑寅卯辰巳午未申酉戌亥") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "cjk-decimal":
                formatted = FormatDecimalDigits(value, "〇一二三四五六七八九");
                return true;
            case "full-width":
                formatted = FormatDecimalDigits(value, "０１２３４５６７８９");
                return true;
            case "japanese-informal":
                formatted = FormatEastAsianAdditive(value, "〇一二三四五六七八九", "十百千", omitOneBeforeUnits: true, "マイナス");
                return true;
            case "japanese-formal":
                formatted = FormatEastAsianAdditive(value, "零壱弐参四伍六七八九", "拾百阡", omitOneBeforeUnits: false, "マイナス");
                return true;
            case "korean-hangul-formal":
                formatted = FormatEastAsianAdditive(value, "영일이삼사오육칠팔구", "십백천", omitOneBeforeUnits: false, "마이너스 ");
                return true;
            case "korean-hanja-informal":
                formatted = FormatEastAsianAdditive(value, "零一二三四五六七八九", "十百千", omitOneBeforeUnits: true, "마이너스 ");
                return true;
            case "korean-hanja-formal":
                formatted = FormatEastAsianAdditive(value, "零壹貳參四五六七八九", "拾百仟", omitOneBeforeUnits: false, "마이너스 ");
                return true;
            case "simp-chinese-informal":
                formatted = FormatChineseLonghand(value, "零一二三四五六七八九", "十百千", informal: true, "负");
                return true;
            case "simp-chinese-formal":
                formatted = FormatChineseLonghand(value, "零壹贰叁肆伍陆柒捌玖", "拾佰仟", informal: false, "负");
                return true;
            case "trad-chinese-informal":
            case "cjk-ideographic":
                formatted = FormatChineseLonghand(value, "零一二三四五六七八九", "十百千", informal: true, "負");
                return true;
            case "trad-chinese-formal":
                formatted = FormatChineseLonghand(value, "零壹貳參肆伍陸柒捌玖", "拾佰仟", informal: false, "負");
                return true;
            case "lower-roman":
                formatted = value > 0 && value <= 3999 ? FormatRoman(value).ToLowerInvariant() : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "upper-roman":
                formatted = value > 0 && value <= 3999 ? FormatRoman(value) : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "disc":
                formatted = "•";
                return true;
            case "circle":
                formatted = "◦";
                return true;
            case "square":
                formatted = "▪";
                return true;
            case "none":
                formatted = string.Empty;
                return true;
            case "decimal":
            case "":
                formatted = value.ToString(CultureInfo.InvariantCulture);
                return true;
            default:
                formatted = string.Empty;
                return false;
        }
    }

    internal static string MarkerSuffix(string style, bool ordered) {
        string normalized = HtmlCssEscapeDecoder.Decode(style.Trim()).ToLowerInvariant();
        if (normalized is "japanese-informal" or "japanese-formal"
            or "simp-chinese-informal" or "simp-chinese-formal"
            or "trad-chinese-informal" or "trad-chinese-formal" or "cjk-ideographic") return "、";
        if (normalized is "korean-hangul-formal" or "korean-hanja-informal" or "korean-hanja-formal") return ", ";
        return ordered ? ". " : " ";
    }

    private static bool TryFormatSymbolsFunction(int value, string style, out string formatted, out bool representationLimited) {
        formatted = string.Empty;
        representationLimited = false;
        const string prefix = "symbols(";
        if (!style.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) || !style.EndsWith(")", StringComparison.Ordinal)) return false;
        string body = style.Substring(prefix.Length, style.Length - prefix.Length - 1).Trim();
        if (!TryTokenizeSymbols(body, out IReadOnlyList<string> tokens) || tokens.Count == 0) return false;
        string system = "symbolic";
        int symbolStart = 0;
        string candidate = tokens[0].ToLowerInvariant();
        if (candidate is "cyclic" or "numeric" or "alphabetic" or "symbolic" or "fixed") {
            system = candidate;
            symbolStart = 1;
        } else if (candidate is "additive") {
            return false;
        }
        var symbols = new List<string>();
        for (int index = symbolStart; index < tokens.Count; index++) {
            if (!TryUnquote(tokens[index], out string symbol) || symbol.Length == 0) return false;
            symbols.Add(symbol);
        }
        if (symbols.Count == 0 || (system is "numeric" or "alphabetic") && symbols.Count < 2) return false;

        switch (system) {
            case "cyclic":
                int cyclicIndex = ((value - 1) % symbols.Count + symbols.Count) % symbols.Count;
                formatted = symbols[cyclicIndex];
                return true;
            case "fixed":
                formatted = value >= 1 && value <= symbols.Count
                    ? symbols[value - 1]
                    : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "numeric":
                formatted = FormatNumericSymbols(value, symbols);
                return true;
            case "alphabetic":
                formatted = value > 0 ? FormatAlphabetic(value, symbols) : value.ToString(CultureInfo.InvariantCulture);
                return true;
            default:
                if (value <= 0) {
                    formatted = value.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
                int symbolicIndex = (value - 1) % symbols.Count;
                int repetitions = ((value - 1) / symbols.Count) + 1;
                if (!TryRepeatSymbol(symbols[symbolicIndex], repetitions, out formatted)) {
                    formatted = value.ToString(CultureInfo.InvariantCulture);
                    representationLimited = true;
                }
                return true;
        }
    }

    internal static bool TryRepeatSymbol(string symbol, int repetitions, out string repeated) {
        repeated = string.Empty;
        if (repetitions < 0 || symbol.Length == 0) return false;
        if (repetitions == 0) return true;
        if (repetitions > MaximumGeneratedRepresentationLength
            || (long)symbol.Length * repetitions > MaximumGeneratedRepresentationLength) return false;
        var builder = new StringBuilder(symbol.Length * repetitions);
        for (int index = 0; index < repetitions; index++) builder.Append(symbol);
        repeated = builder.ToString();
        return true;
    }

    internal static bool TryTokenizeSymbols(string value, out IReadOnlyList<string> tokens) {
        var result = new List<string>();
        int cursor = 0;
        while (cursor < value.Length) {
            while (cursor < value.Length && char.IsWhiteSpace(value[cursor])) cursor++;
            if (cursor >= value.Length) break;
            int start = cursor;
            if (value[cursor] == '\'' || value[cursor] == '"') {
                char quote = value[cursor++];
                while (cursor < value.Length) {
                    if (value[cursor] == quote && !IsEscaped(value, cursor)) {
                        cursor++;
                        break;
                    }
                    cursor++;
                }
                if (cursor > value.Length || value[cursor - 1] != quote) {
                    tokens = Array.Empty<string>();
                    return false;
                }
            } else {
                while (cursor < value.Length && !char.IsWhiteSpace(value[cursor])) cursor++;
            }
            result.Add(value.Substring(start, cursor - start));
        }
        tokens = result.AsReadOnly();
        return true;
    }

    private static bool IsEscaped(string value, int index) {
        int slashes = 0;
        for (int cursor = index - 1; cursor >= 0 && value[cursor] == '\\'; cursor--) slashes++;
        return (slashes & 1) != 0;
    }

    internal static string FormatNumericSymbols(int value, IReadOnlyList<string> symbols) {
        if (value == 0) return symbols[0];
        bool negative = value < 0;
        long remaining = Math.Abs((long)value);
        var parts = new List<string>();
        while (remaining > 0) {
            parts.Add(symbols[(int)(remaining % symbols.Count)]);
            remaining /= symbols.Count;
        }
        parts.Reverse();
        return (negative ? "-" : string.Empty) + string.Concat(parts);
    }

    private static string FormatDecimalDigits(int value, string digits) {
        string source = Math.Abs((long)value).ToString(CultureInfo.InvariantCulture);
        var result = new StringBuilder(source.Length + 1);
        if (value < 0) result.Append('-');
        foreach (char digit in source) result.Append(digits[digit - '0']);
        return result.ToString();
    }

    private static string FormatEastAsianAdditive(
        int value,
        string digits,
        string units,
        bool omitOneBeforeUnits,
        string negativePrefix) {
        if (value < -9999 || value > 9999) return FormatDecimalDigits(value, "〇一二三四五六七八九");
        if (value == 0) return digits[0].ToString();
        bool negative = value < 0;
        int magnitude = Math.Abs(value);
        var result = new StringBuilder();
        int[] divisors = { 1000, 100, 10 };
        for (int index = 0; index < divisors.Length; index++) {
            int digit = magnitude / divisors[index];
            magnitude %= divisors[index];
            if (digit == 0) continue;
            if (!omitOneBeforeUnits || digit != 1) result.Append(digits[digit]);
            result.Append(units[2 - index]);
        }
        if (magnitude > 0) result.Append(digits[magnitude]);
        return (negative ? negativePrefix : string.Empty) + result;
    }

    private static string FormatChineseLonghand(
        int value,
        string digits,
        string units,
        bool informal,
        string negativePrefix) {
        if (value < -9999 || value > 9999) return FormatDecimalDigits(value, "〇一二三四五六七八九");
        if (value == 0) return digits[0].ToString();
        bool negative = value < 0;
        int magnitude = Math.Abs(value);
        var result = new StringBuilder();
        bool emitted = false;
        bool zeroPending = false;
        int[] divisors = { 1000, 100, 10, 1 };
        for (int index = 0; index < divisors.Length; index++) {
            int digit = magnitude / divisors[index];
            magnitude %= divisors[index];
            if (digit == 0) {
                if (emitted && magnitude > 0) zeroPending = true;
                continue;
            }
            if (zeroPending) {
                result.Append(digits[0]);
                zeroPending = false;
            }
            bool omitLeadingOne = informal && index == 2 && digit == 1 && !emitted;
            if (!omitLeadingOne) result.Append(digits[digit]);
            if (index < 3) result.Append(units[2 - index]);
            emitted = true;
        }
        return (negative ? negativePrefix : string.Empty) + result;
    }

    internal static bool TryUnquote(string value, out string result) {
        if (value.Length >= 2 && (value[0] == '\'' && value[value.Length - 1] == '\'' || value[0] == '"' && value[value.Length - 1] == '"')) {
            result = value.Substring(1, value.Length - 2);
            return true;
        }
        result = string.Empty;
        return false;
    }

    private static string FormatAlphabetic(int value, string alphabet) {
        var result = new StringBuilder();
        int remaining = value;
        while (remaining > 0) {
            remaining--;
            result.Insert(0, alphabet[remaining % alphabet.Length]);
            remaining /= alphabet.Length;
        }
        return result.ToString();
    }

    internal static string FormatAlphabetic(int value, IReadOnlyList<string> alphabet) {
        var result = new List<string>();
        int remaining = value;
        while (remaining > 0) {
            remaining--;
            result.Add(alphabet[remaining % alphabet.Count]);
            remaining /= alphabet.Count;
        }
        result.Reverse();
        return string.Concat(result);
    }

    private static string FormatRoman(int value) {
        var result = new StringBuilder();
        int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        for (int index = 0; index < values.Length; index++) {
            while (value >= values[index]) {
                result.Append(symbols[index]);
                value -= values[index];
            }
        }
        return result.ToString();
    }
}
