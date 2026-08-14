using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string ManagedConicGradientUrlPrefix = "officeimo-managed-conic:";
    private static readonly Regex ManagedConicGradientUrl = new Regex(
        @"url\(\s*['\""']?officeimo-managed-conic:(?<payload>[A-Za-z0-9+/=]+)['\""']?\s*\)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static string PreserveManagedGradientFunctions(string css) {
        if (string.IsNullOrEmpty(css)
            || css.IndexOf("conic-gradient(", StringComparison.OrdinalIgnoreCase) < 0) {
            return css;
        }

        var result = new StringBuilder(css.Length);
        char quote = '\0';
        for (int index = 0; index < css.Length;) {
            char current = css[index];
            if (quote != '\0') {
                result.Append(current);
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                index++;
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                result.Append(current);
                index++;
                continue;
            }
            if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                int commentClose = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                int commentEnd = commentClose < 0 ? css.Length : commentClose + 2;
                result.Append(css, index, commentEnd - index);
                index = commentEnd;
                continue;
            }
            if (TryGetManagedGradientFunctionLength(css, index, out int functionNameLength)
                && TryFindFunctionEnd(css, index + functionNameLength - 1, out int functionEnd)) {
                string function = css.Substring(index, functionEnd - index + 1);
                string payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(function));
                result.Append("url(\"").Append(ManagedConicGradientUrlPrefix).Append(payload).Append("\")");
                index = functionEnd + 1;
                continue;
            }
            result.Append(current);
            index++;
        }
        return result.ToString();
    }

    private static string RestoreManagedGradientFunctions(string value) {
        if (string.IsNullOrEmpty(value)
            || value.IndexOf(ManagedConicGradientUrlPrefix, StringComparison.OrdinalIgnoreCase) < 0) {
            return value;
        }

        return ManagedConicGradientUrl.Replace(value, match => {
            try {
                byte[] bytes = Convert.FromBase64String(match.Groups["payload"].Value);
                string function = Encoding.UTF8.GetString(bytes);
                return TryGetManagedGradientFunctionLength(function, 0, out int functionNameLength)
                    && TryFindFunctionEnd(function, functionNameLength - 1, out int functionEnd)
                    && functionEnd == function.Length - 1
                        ? function
                        : match.Value;
            } catch (FormatException) {
                return match.Value;
            }
        });
    }

    private static bool TryGetManagedGradientFunctionLength(string css, int index, out int length) {
        length = 0;
        if (index > 0 && IsCssIdentifierCharacter(css[index - 1])) return false;
        const string repeating = "repeating-conic-gradient(";
        const string conic = "conic-gradient(";
        if (MatchesAt(css, index, repeating)) length = repeating.Length;
        else if (MatchesAt(css, index, conic)) length = conic.Length;
        else return false;
        return true;
    }

    private static bool MatchesAt(string value, int index, string token) =>
        index + token.Length <= value.Length
        && string.Compare(value, index, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) == 0;

    private static bool IsCssIdentifierCharacter(char value) =>
        char.IsLetterOrDigit(value) || value == '-' || value == '_' || value >= 0x80;

    private static bool TryFindFunctionEnd(string css, int openParenthesis, out int end) {
        end = -1;
        int depth = 0;
        char quote = '\0';
        for (int index = openParenthesis; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                continue;
            }
            if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                int commentClose = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (commentClose < 0) return false;
                index = commentClose + 1;
                continue;
            }
            if (current == '(') depth++;
            else if (current == ')' && --depth == 0) {
                end = index;
                return true;
            }
        }
        return false;
    }
}
