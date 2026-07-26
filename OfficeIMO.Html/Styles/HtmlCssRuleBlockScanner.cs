using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Lexically scans CSS rule blocks before handing untrusted stylesheets to a parser.
/// </summary>
internal static class HtmlCssRuleBlockScanner {
    internal static IReadOnlyDictionary<int, int> Scan(
        string css,
        HtmlCssProcessingBudget budget) {
        var closures = new Dictionary<int, int>();
        var opens = new Stack<int>();
        char quote = '\0';
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') quote = current;
            else if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                index = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (index < 0) break;
                index++;
            } else if (current == '{' && !IsEscaped(css, index)) {
                opens.Push(index);
                budget.RecordNestingDepth(opens.Count);
            } else if (current == '}' && !IsEscaped(css, index) && opens.Count > 0) {
                closures[opens.Pop()] = index;
            }
        }
        return closures;
    }

    internal static void ValidateDocument(
        IHtmlDocument document,
        HtmlConversionLimits limits) {
        var budget = new HtmlCssProcessingBudget(limits);
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement)) continue;
            string css = styleElement.TextContent;
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
        string type = (styleElement.GetAttribute("type") ?? string.Empty).Trim();
        int parameterStart = type.IndexOf(';');
        if (parameterStart >= 0) type = type.Substring(0, parameterStart).Trim();
        return type.Length == 0
            || string.Equals(type, "text/css", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEscaped(string text, int index) {
        int backslashes = 0;
        for (int cursor = index - 1; cursor >= 0 && text[cursor] == '\\'; cursor--) {
            backslashes++;
        }
        return (backslashes & 1) != 0;
    }
}
