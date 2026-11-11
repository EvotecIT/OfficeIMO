using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Lexically scans CSS rule blocks before handing untrusted stylesheets to a parser.
/// </summary>
internal static class HtmlCssRuleBlockScanner {
    internal static IReadOnlyDictionary<int, int> Scan(
        string css,
        HtmlCssProcessingBudget budget) => Scan(css, budget, out _);

    /// <summary>Scans <paramref name="css"/>, also reporting the deepest block nesting it recorded.</summary>
    internal static IReadOnlyDictionary<int, int> Scan(
        string css,
        HtmlCssProcessingBudget budget,
        out int maximumDepth) {
        maximumDepth = 0;
        var closures = new Dictionary<int, int>();
        var opens = new Stack<int>();
        char quote = '\0';
        bool insideUrl = false;
        int identifierLength = 0;
        bool identifierMatchesUrl = true;
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (IsCssNewline(current) && !IsEscapedCssNewline(css, index)) {
                    quote = '\0';
                } else if (current == quote && !IsEscaped(css, index)) {
                    quote = '\0';
                }
                continue;
            }
            if (current == '\'' || current == '"') {
                // A quoted url() is a function containing a string, not an unquoted URL token.
                insideUrl = false;
                quote = current;
                ResetIdentifier(ref identifierLength, ref identifierMatchesUrl);
            } else if (!insideUrl && current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                // Comments are token boundaries, but slash/star within a URL token is path data.
                ResetIdentifier(ref identifierLength, ref identifierMatchesUrl);
                index = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (index < 0) break;
                index++;
            } else if (current == '\\'
                && HtmlCssEscapeDecoder.TryDecodeEscape(
                    css,
                    index,
                    out string decoded,
                    out int consumedCharacters)) {
                if (!insideUrl) {
                    AppendIdentifier(decoded, ref identifierLength, ref identifierMatchesUrl);
                }
                index += consumedCharacters - 1;
            } else if (insideUrl) {
                if (current == ')') insideUrl = false;
            } else if (IsIdentifierCharacter(current)) {
                AppendIdentifier(current, ref identifierLength, ref identifierMatchesUrl);
            } else if (current == '(') {
                insideUrl = identifierMatchesUrl && identifierLength == 3;
                ResetIdentifier(ref identifierLength, ref identifierMatchesUrl);
            } else if (current == '{') {
                ResetIdentifier(ref identifierLength, ref identifierMatchesUrl);
                opens.Push(index);
                budget.RecordNestingDepth(opens.Count);
                if (opens.Count > maximumDepth) maximumDepth = opens.Count;
            } else {
                ResetIdentifier(ref identifierLength, ref identifierMatchesUrl);
                if (current == '}' && opens.Count > 0) {
                    closures[opens.Pop()] = index;
                }
            }
        }
        return closures;
    }

    /// <summary>Whether the block opened at <paramref name="open"/> is introduced by a literal <c>@font-face</c> keyword.</summary>
    internal static bool IsFontFaceBlock(string css, int open) {
        int cursor = open - 1;
        while (cursor >= 0 && char.IsWhiteSpace(css[cursor])) cursor--;
        const string RuleName = "@font-face";
        int ruleStart = cursor - RuleName.Length + 1;
        return ruleStart >= 0
            && string.Compare(css, ruleStart, RuleName, 0, RuleName.Length, StringComparison.OrdinalIgnoreCase) == 0
            && !(ruleStart > 0 && (char.IsLetterOrDigit(css[ruleStart - 1]) || css[ruleStart - 1] == '-' || css[ruleStart - 1] == '_'));
    }

    internal static void ValidateDocument(
        IHtmlDocument document,
        HtmlConversionLimits limits) {
        var budget = new HtmlCssProcessingBudget(limits);
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement)) continue;
            string css = HtmlStyleElementText.Read(styleElement);
            if (!string.IsNullOrWhiteSpace(css)) {
                Scan(css, budget);
            }
        }
    }

    internal static void ValidateStylesheet(
        string css,
        HtmlConversionLimits? limits) {
        if (!string.IsNullOrWhiteSpace(css)) {
            Scan(css, new HtmlCssProcessingBudget(limits));
        }
    }

    private static bool IsCssStyleElement(IElement styleElement) {
        return HtmlResourcePipeline.IsCssStyleElement(styleElement);
    }

    private static bool IsEscaped(string text, int index) {
        int backslashes = 0;
        for (int cursor = index - 1; cursor >= 0 && text[cursor] == '\\'; cursor--) {
            backslashes++;
        }
        return (backslashes & 1) != 0;
    }

    private static bool IsEscapedCssNewline(string text, int index) =>
        IsEscaped(text, index)
        || (text[index] == '\n'
            && index > 0
            && text[index - 1] == '\r'
            && IsEscaped(text, index - 1));

    private static bool IsCssNewline(char value) =>
        value == '\n' || value == '\r' || value == '\f';

    private static bool IsIdentifierCharacter(char value) =>
        char.IsLetterOrDigit(value) || value == '_' || value == '-' || value >= 0x80;

    private static void AppendIdentifier(
        string value,
        ref int identifierLength,
        ref bool identifierMatchesUrl) {
        for (int index = 0; index < value.Length; index++) {
            AppendIdentifier(value[index], ref identifierLength, ref identifierMatchesUrl);
        }
    }

    private static void AppendIdentifier(
        char value,
        ref int identifierLength,
        ref bool identifierMatchesUrl) {
        const string Url = "url";
        if (identifierLength >= Url.Length
            || char.ToLowerInvariant(value) != Url[identifierLength]) {
            identifierMatchesUrl = false;
        }
        identifierLength++;
    }

    private static void ResetIdentifier(
        ref int identifierLength,
        ref bool identifierMatchesUrl) {
        identifierLength = 0;
        identifierMatchesUrl = true;
    }
}
