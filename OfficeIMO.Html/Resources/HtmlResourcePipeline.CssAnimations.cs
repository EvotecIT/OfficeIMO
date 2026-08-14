using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static bool ContainsAnimationLonghandName(string value, string name) {
        if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(name)) return false;
        foreach (string item in SplitTopLevelList(value)) {
            string candidate = DecodeAnimationName(item);
            if (string.Equals(candidate, name, StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static bool ContainsAnimationShorthandName(string value, string name) {
        if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(name)) return false;
        foreach (string item in SplitTopLevelList(value)) {
            string? candidate = GetAnimationShorthandName(item);
            if (string.Equals(candidate, name, StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static string? GetAnimationShorthandName(string item) {
        string[] tokens = SplitAnimationTokens(item).ToArray();
        if (tokens.Length == 0) return null;
        if (tokens.Length == 1 && IsCssWideKeyword(DecodeCssEscapes(tokens[0]).Trim())) return null;

        bool duration = false;
        bool delay = false;
        bool timing = false;
        bool iteration = false;
        bool direction = false;
        bool fill = false;
        bool play = false;
        string? name = null;
        foreach (string rawToken in tokens) {
            string token = DecodeCssEscapes(rawToken).Trim();
            bool quotedName = rawToken.Length >= 2 &&
                rawToken[0] == rawToken[rawToken.Length - 1] &&
                rawToken[0] is '\'' or '"';
            if (quotedName) {
                if (name != null) return null;
                name = DecodeAnimationName(rawToken);
                continue;
            }
            string keyword = token.ToLowerInvariant();
            if (IsAnimationTime(keyword)) {
                if (!duration) duration = true;
                else if (!delay) delay = true;
                else return null;
                continue;
            }
            if (IsAnimationTimingFunction(keyword)) {
                if (!timing) {
                    timing = true;
                    continue;
                }
                if (keyword.IndexOf('(') >= 0) return null;
            }
            if (IsAnimationIterationCount(keyword)) {
                if (!iteration) {
                    iteration = true;
                    continue;
                }
                if (keyword != "infinite") return null;
            }
            if (keyword is "normal" or "reverse" or "alternate" or "alternate-reverse") {
                if (!direction) {
                    direction = true;
                    continue;
                }
            }
            if (keyword is "none" or "forwards" or "backwards" or "both") {
                if (!fill) {
                    fill = true;
                    continue;
                }
                if (keyword == "none") return null;
            }
            if (keyword is "running" or "paused") {
                if (!play) {
                    play = true;
                    continue;
                }
            }
            if (IsCssWideKeyword(keyword) || name != null) return null;
            name = DecodeAnimationName(token);
        }
        return name;
    }

    private static IEnumerable<string> SplitAnimationTokens(string value) {
        var token = new StringBuilder();
        int parentheses = 0;
        char quote = '\0';
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current == '\\') {
                token.Append(current);
                if (index + 1 < value.Length) token.Append(value[++index]);
                continue;
            }
            if (quote != '\0') {
                token.Append(current);
                if (current == quote) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') {
                quote = current;
                token.Append(current);
                continue;
            }
            if (current == '(') {
                parentheses++;
                token.Append(current);
                continue;
            }
            if (current == ')') {
                parentheses = Math.Max(0, parentheses - 1);
                token.Append(current);
                continue;
            }
            if (parentheses == 0 && IsCssWhitespace(current)) {
                if (token.Length != 0) {
                    yield return token.ToString();
                    token.Clear();
                }
                continue;
            }
            token.Append(current);
        }
        if (quote == '\0' && parentheses == 0 && token.Length != 0) yield return token.ToString();
    }

    private static string DecodeAnimationName(string value) {
        string trimmed = value.Trim();
        if (trimmed.Length >= 2 && trimmed[0] == trimmed[trimmed.Length - 1] && trimmed[0] is '\'' or '"') {
            trimmed = trimmed.Substring(1, trimmed.Length - 2);
        }
        return DecodeCssEscapes(trimmed);
    }

    private static bool IsAnimationTime(string token) {
        string number;
        if (token.EndsWith("ms", StringComparison.Ordinal)) number = token.Substring(0, token.Length - 2);
        else if (token.EndsWith("s", StringComparison.Ordinal)) number = token.Substring(0, token.Length - 1);
        else return false;
        return double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
    }

    private static bool IsAnimationTimingFunction(string token) =>
        token is "ease" or "ease-in" or "ease-out" or "ease-in-out" or "linear" or "step-start" or "step-end" ||
        token.StartsWith("cubic-bezier(", StringComparison.Ordinal) ||
        token.StartsWith("steps(", StringComparison.Ordinal) ||
        token.StartsWith("linear(", StringComparison.Ordinal);

    private static bool IsAnimationIterationCount(string token) =>
        token == "infinite" || double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) && value >= 0;

    private static bool IsCssWideKeyword(string token) =>
        token is "inherit" or "initial" or "revert" or "revert-layer" or "unset";
}
