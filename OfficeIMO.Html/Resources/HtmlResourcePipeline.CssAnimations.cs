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
            TryParseAnimationShorthand(item, out string? candidate);
            if (string.Equals(candidate, name, StringComparison.Ordinal)) return true;
        }
        return false;
    }

    internal static bool TryExpandAnimationShorthandNames(string value, out string names) {
        if (HtmlCssCustomPropertyResolver.ContainsVarFunction(value)) {
            names = string.Empty;
            return false;
        }
        var expanded = new List<string>();
        string[] items = SplitTopLevelList(value).ToArray();
        foreach (string item in items) {
            string normalized = DecodeCssEscapes(item).Trim();
            if (IsCssWideKeyword(normalized.ToLowerInvariant())) {
                if (items.Length != 1) {
                    names = string.Empty;
                    return false;
                }
                expanded.Add(normalized);
                continue;
            }
            if (!TryParseAnimationShorthand(item, out string? name)) {
                names = string.Empty;
                return false;
            }
            expanded.Add(name ?? "none");
        }
        names = string.Join(", ", expanded);
        return expanded.Count != 0;
    }

    private static bool TryParseAnimationShorthand(string item, out string? name) {
        string[] tokens = SplitAnimationTokens(item).ToArray();
        name = null;
        if (tokens.Length == 0) return false;
        if (tokens.Length == 1 && IsCssWideKeyword(DecodeCssEscapes(tokens[0]).Trim())) return true;

        bool duration = false;
        bool delay = false;
        bool timing = false;
        bool iteration = false;
        bool direction = false;
        bool fill = false;
        bool play = false;
        foreach (string rawToken in tokens) {
            string token = DecodeCssEscapes(rawToken).Trim();
            bool quotedName = rawToken.Length >= 2 &&
                rawToken[0] == rawToken[rawToken.Length - 1] &&
                rawToken[0] is '\'' or '"';
            if (quotedName) {
                if (name != null) return false;
                name = DecodeAnimationName(rawToken);
                continue;
            }
            string keyword = token.ToLowerInvariant();
            if (TryReadAnimationTime(keyword, out double seconds)) {
                if (!duration) {
                    if (seconds < 0D) return false;
                    duration = true;
                }
                else if (!delay) delay = true;
                else return false;
                continue;
            }
            if (IsAnimationTimingFunction(keyword)) {
                if (!timing) {
                    timing = true;
                    continue;
                }
                if (LooksLikeAnimationTimingFunction(keyword)) return false;
            }
            else if (LooksLikeAnimationTimingFunction(keyword)) return false;
            if (IsAnimationIterationCount(keyword)) {
                if (!iteration) {
                    iteration = true;
                    continue;
                }
                if (keyword != "infinite") return false;
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
                if (keyword == "none") return false;
            }
            if (keyword is "running" or "paused") {
                if (!play) {
                    play = true;
                    continue;
                }
            }
            if (IsCssWideKeyword(keyword) || name != null) return false;
            name = DecodeAnimationName(token);
        }
        return true;
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

    private static bool TryReadAnimationTime(string token, out double seconds) {
        string number;
        double divisor;
        if (token.EndsWith("ms", StringComparison.Ordinal)) {
            number = token.Substring(0, token.Length - 2);
            divisor = 1000D;
        } else if (token.EndsWith("s", StringComparison.Ordinal)) {
            number = token.Substring(0, token.Length - 1);
            divisor = 1D;
        } else {
            seconds = 0D;
            return false;
        }
        if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
            double.IsNaN(parsed) || double.IsInfinity(parsed)) {
            seconds = 0D;
            return false;
        }
        seconds = parsed / divisor;
        return true;
    }

    private static bool IsAnimationTimingFunction(string token) {
        if (token is "ease" or "ease-in" or "ease-out" or "ease-in-out" or "linear" or "step-start" or "step-end") return true;
        if (!TryReadFunctionArguments(token, "cubic-bezier", out string[] bezier) || bezier.Length != 4) {
            if (!TryReadFunctionArguments(token, "steps", out string[] steps)) {
                return TryReadFunctionArguments(token, "linear", out string[] stops) && IsValidLinearStops(stops);
            }
            if (steps.Length is < 1 or > 2 || !int.TryParse(steps[0], NumberStyles.None, CultureInfo.InvariantCulture, out int intervals) || intervals <= 0) return false;
            if (steps.Length == 1) return true;
            string position = steps[1].Trim();
            if (position is not ("jump-start" or "jump-end" or "jump-none" or "jump-both" or "start" or "end")) return false;
            return position != "jump-none" || intervals >= 2;
        }
        return TryReadFiniteNumber(bezier[0], out double x1) && x1 is >= 0D and <= 1D &&
            TryReadFiniteNumber(bezier[1], out _) &&
            TryReadFiniteNumber(bezier[2], out double x2) && x2 is >= 0D and <= 1D &&
            TryReadFiniteNumber(bezier[3], out _);
    }

    private static bool TryReadFunctionArguments(string token, string name, out string[] arguments) {
        arguments = Array.Empty<string>();
        string prefix = name + "(";
        if (!token.StartsWith(prefix, StringComparison.Ordinal) || token.Length <= prefix.Length || token[token.Length - 1] != ')') return false;
        string body = token.Substring(prefix.Length, token.Length - prefix.Length - 1);
        if (body.IndexOf('(') >= 0 || body.IndexOf(')') >= 0) return false;
        arguments = body.Split(',').Select(static item => item.Trim()).ToArray();
        return arguments.Length != 0 && arguments.All(static item => item.Length != 0);
    }

    private static bool IsValidLinearStops(string[] stops) {
        if (stops.Length < 2) return false;
        foreach (string stop in stops) {
            string[] components = stop.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (components.Length is < 1 or > 3 || !TryReadFiniteNumber(components[0], out _)) return false;
            for (int index = 1; index < components.Length; index++) {
                string position = components[index];
                if (!position.EndsWith("%", StringComparison.Ordinal) ||
                    !TryReadFiniteNumber(position.Substring(0, position.Length - 1), out _)) return false;
            }
        }
        return true;
    }

    private static bool TryReadFiniteNumber(string token, out double value) =>
        double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
        !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool LooksLikeAnimationTimingFunction(string token) =>
        token.StartsWith("cubic-bezier(", StringComparison.Ordinal) ||
        token.StartsWith("steps(", StringComparison.Ordinal) ||
        token.StartsWith("linear(", StringComparison.Ordinal);

    private static bool IsAnimationIterationCount(string token) =>
        token == "infinite" ||
        double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
        !double.IsNaN(value) && !double.IsInfinity(value) && value >= 0;

    private static bool IsCssWideKeyword(string token) =>
        token is "inherit" or "initial" or "revert" or "revert-layer" or "unset";
}
