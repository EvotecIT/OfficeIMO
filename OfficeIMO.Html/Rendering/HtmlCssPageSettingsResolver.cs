using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssPageSettingsResolver {
    internal static void Apply(IHtmlDocument document, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        if (options.Mode != HtmlRenderMode.Paged || !options.HonorCssPageRules) return;
        var parser = new CssParser();
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !HtmlComputedStyleEngine.IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, HtmlCssMediaContext.Print)) continue;
            ApplyRawGenericPageSizes(styleElement.TextContent, options, diagnostics);
            var sheet = parser.ParseStyleSheet(styleElement.TextContent);
            foreach (ICssRule rule in sheet.Rules) ApplyRule(rule, options, diagnostics);
        }
    }

    private static void ApplyRule(ICssRule rule, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        if (rule is ICssMediaRule mediaRule && !HtmlComputedStyleEngine.IsApplicableMedia(mediaRule.ConditionText, HtmlCssMediaContext.Print)) return;
        if (rule is ICssPageRule pageRule) {
            string selector = (pageRule.SelectorText ?? string.Empty).Trim();
            if (selector.Length > 0) {
                diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSelectorPending, "A named or pseudo-page rule is not yet applied to individual pages.", HtmlDiagnosticSeverity.Warning, selector);
                return;
            }

            ApplyPageRule(pageRule, options, diagnostics);
            return;
        }

        if (rule is ICssGroupingRule grouping) {
            foreach (ICssRule child in grouping.Rules) ApplyRule(child, options, diagnostics);
        }
    }

    private static void ApplyPageRule(ICssPageRule pageRule, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        string size = pageRule.Style.GetPropertyValue("size");
        if (!string.IsNullOrWhiteSpace(size) && !TryApplyPageSize(size, options)) {
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSizeUnsupported, "The @page size declaration could not be mapped to a supported physical page size.", HtmlDiagnosticSeverity.Warning, "@page", size);
        }

        double top = options.Margins.Top;
        double right = options.Margins.Right;
        double bottom = options.Margins.Bottom;
        double left = options.Margins.Left;
        string margin = pageRule.Style.GetPropertyValue("margin");
        if (!string.IsNullOrWhiteSpace(margin)) HtmlRenderCssValues.ApplyBoxShorthand(margin, options.PageWidth, options.DefaultFontSize, options.DefaultFontSize, ref top, ref right, ref bottom, ref left);
        ApplyMarginSide(pageRule.Style.GetPropertyValue("margin-top"), options, ref top);
        ApplyMarginSide(pageRule.Style.GetPropertyValue("margin-right"), options, ref right);
        ApplyMarginSide(pageRule.Style.GetPropertyValue("margin-bottom"), options, ref bottom);
        ApplyMarginSide(pageRule.Style.GetPropertyValue("margin-left"), options, ref left);
        options.Margins = new HtmlRenderMargins(left, top, right, bottom);
    }

    private static bool TryApplyPageSize(string value, HtmlRenderOptions options) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        if (parts.Count == 0) return false;
        OfficePageSize? named = ResolveNamedSize(parts[0]);
        bool landscape = parts.Any(part => string.Equals(part, "landscape", StringComparison.OrdinalIgnoreCase));
        bool portrait = parts.Any(part => string.Equals(part, "portrait", StringComparison.OrdinalIgnoreCase));
        if (named.HasValue) {
            options.PageSize = landscape ? named.Value.Landscape() : portrait ? named.Value.Portrait() : named.Value;
            return true;
        }

        var lengths = new List<double>();
        foreach (string part in parts) {
            if (string.Equals(part, "landscape", StringComparison.OrdinalIgnoreCase) || string.Equals(part, "portrait", StringComparison.OrdinalIgnoreCase)) continue;
            if (!HtmlRenderCssValues.TryLength(part, options.PageWidth, options.DefaultFontSize, options.DefaultFontSize, out double length) || length <= 0D) return false;
            lengths.Add(length);
        }

        if (lengths.Count != 2) return false;
        var custom = new OfficePageSize(lengths[0] / HtmlRenderOptions.CssPixelsPerInch, lengths[1] / HtmlRenderOptions.CssPixelsPerInch);
        options.PageSize = landscape ? custom.Landscape() : portrait ? custom.Portrait() : custom;
        return true;
    }

    private static OfficePageSize? ResolveNamedSize(string value) {
        switch (value.Trim().ToLowerInvariant()) {
            case "a3": return OfficePageSizes.A3;
            case "a4": return OfficePageSizes.A4;
            case "a5": return OfficePageSizes.A5;
            case "b4": return OfficePageSizes.B4Jis;
            case "b5": return OfficePageSizes.B5Jis;
            case "letter": return OfficePageSizes.Letter;
            case "legal": return OfficePageSizes.Legal;
            case "ledger": return OfficePageSizes.Ledger;
            case "tabloid": return OfficePageSizes.Tabloid;
            case "statement": return OfficePageSizes.Statement;
            case "executive": return OfficePageSizes.Executive;
            default: return null;
        }
    }

    private static void ApplyMarginSide(string value, HtmlRenderOptions options, ref double target) {
        if (HtmlRenderCssValues.TryLength(value, options.PageWidth, options.DefaultFontSize, options.DefaultFontSize, out double parsed)) target = Math.Max(0D, parsed);
    }

    private static bool IsCssStyleElement(IElement element) {
        string type = (element.GetAttribute("type") ?? string.Empty).Trim();
        int separator = type.IndexOf(';');
        if (separator >= 0) type = type.Substring(0, separator).Trim();
        return type.Length == 0 || string.Equals(type, "text/css", StringComparison.OrdinalIgnoreCase);
    }

    private static void ApplyRawGenericPageSizes(string css, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) =>
        ScanRawRules(css, 0, css.Length, options, diagnostics);

    private static void ScanRawRules(string css, int start, int end, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        int cursor = start;
        while (cursor < end) {
            if (IsCommentStart(css, cursor)) {
                cursor = SkipComment(css, cursor + 2, end);
                continue;
            }

            char current = css[cursor];
            if (current == '\'' || current == '"') {
                cursor = SkipQuoted(css, cursor + 1, end, current);
                continue;
            }

            if (current == '{') {
                int close = FindMatchingBrace(css, cursor);
                cursor = close < 0 ? end : close + 1;
                continue;
            }

            if (current != '@') {
                cursor++;
                continue;
            }

            int nameStart = cursor + 1;
            int nameEnd = nameStart;
            while (nameEnd < end && (char.IsLetter(css[nameEnd]) || css[nameEnd] == '-')) nameEnd++;
            string name = css.Substring(nameStart, nameEnd - nameStart);
            int boundary = FindRuleBoundary(css, nameEnd, end);
            if (boundary < 0 || css[boundary] == ';') {
                cursor = boundary < 0 ? end : boundary + 1;
                continue;
            }

            int closeBrace = FindMatchingBrace(css, boundary);
            if (closeBrace < 0 || closeBrace >= end) return;
            string prelude = css.Substring(nameEnd, boundary - nameEnd).Trim();
            if (string.Equals(name, "media", StringComparison.OrdinalIgnoreCase)) {
                if (HtmlComputedStyleEngine.IsApplicableMedia(prelude, HtmlCssMediaContext.Print)) {
                    ScanRawRules(css, boundary + 1, closeBrace, options, diagnostics);
                }
            } else if (string.Equals(name, "page", StringComparison.OrdinalIgnoreCase) && prelude.Length == 0) {
                string body = css.Substring(boundary + 1, closeBrace - boundary - 1);
                string size = FindTopLevelDeclaration(body, "size");
                if (size.Length > 0 && !TryApplyPageSize(size, options)) {
                    diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSizeUnsupported, "The @page size declaration could not be mapped to a supported physical page size.", HtmlDiagnosticSeverity.Warning, "@page", size);
                }
            }

            cursor = closeBrace + 1;
        }
    }

    private static int FindRuleBoundary(string css, int start, int end) {
        int parentheses = 0;
        for (int index = start; index < end; index++) {
            if (IsCommentStart(css, index)) {
                index = SkipComment(css, index + 2, end) - 1;
                continue;
            }

            char current = css[index];
            if (current == '\'' || current == '"') {
                index = SkipQuoted(css, index + 1, end, current) - 1;
            } else if (current == '(') {
                parentheses++;
            } else if (current == ')' && parentheses > 0) {
                parentheses--;
            } else if (parentheses == 0 && (current == '{' || current == ';')) {
                return index;
            }
        }

        return -1;
    }

    private static string FindTopLevelDeclaration(string body, string propertyName) {
        int start = 0;
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index <= body.Length; index++) {
            char current = index < body.Length ? body[index] : ';';
            if (quote != '\0') {
                if (current == quote && (index == 0 || body[index - 1] != '\\')) quote = '\0';
                continue;
            }

            if (current == '\'' || current == '"') quote = current;
            else if (current == '(' || current == '{') depth++;
            else if ((current == ')' || current == '}') && depth > 0) depth--;
            else if (current == ';' && depth == 0) {
                string declaration = body.Substring(start, index - start).Trim();
                int separator = declaration.IndexOf(':');
                if (separator > 0 && string.Equals(declaration.Substring(0, separator).Trim(), propertyName, StringComparison.OrdinalIgnoreCase)) {
                    return declaration.Substring(separator + 1).Trim();
                }

                start = index + 1;
            }
        }

        return string.Empty;
    }

    private static int FindMatchingBrace(string css, int open) {
        int depth = 0;
        char quote = '\0';
        for (int index = open; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && (index == 0 || css[index - 1] != '\\')) quote = '\0';
                continue;
            }

            if (IsCommentStart(css, index)) {
                index = SkipComment(css, index + 2, css.Length) - 1;
            } else if (current == '\'' || current == '"') quote = current;
            else if (current == '{') depth++;
            else if (current == '}' && --depth == 0) return index;
        }

        return -1;
    }

    private static bool IsCommentStart(string css, int index) =>
        index + 1 < css.Length && css[index] == '/' && css[index + 1] == '*';

    private static int SkipComment(string css, int start, int end) {
        int close = css.IndexOf("*/", start, StringComparison.Ordinal);
        return close < 0 || close + 2 > end ? end : close + 2;
    }

    private static int SkipQuoted(string css, int start, int end, char quote) {
        for (int index = start; index < end; index++) {
            if (css[index] == quote && css[index - 1] != '\\') return index + 1;
        }

        return end;
    }
}
