using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssPageSettingsResolver {
    internal static HtmlCssPageRuleSet Apply(IHtmlDocument document, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        var pageRules = new HtmlCssPageRuleSet();
        if (options.Mode != HtmlRenderMode.Paged || !options.HonorCssPageRules) return pageRules;
        var parser = new CssParser();
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !HtmlComputedStyleEngine.IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, HtmlCssMediaContext.Print)) continue;
            ApplyRawPageRules(styleElement.TextContent, options, diagnostics, pageRules);
            var sheet = parser.ParseStyleSheet(styleElement.TextContent);
            foreach (ICssRule rule in sheet.Rules) ApplyRule(rule, options, diagnostics);
        }

        return pageRules;
    }

    private static void ApplyRule(ICssRule rule, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        if (rule is ICssMediaRule mediaRule && !HtmlComputedStyleEngine.IsApplicableMedia(mediaRule.ConditionText, HtmlCssMediaContext.Print)) return;
        if (rule is ICssPageRule pageRule) {
            string selector = (pageRule.SelectorText ?? string.Empty).Trim();
            if (selector.Length > 0) {
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

    private static void ApplyRawPageRules(string css, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, HtmlCssPageRuleSet pageRules) =>
        ScanRawRules(css, 0, css.Length, options, diagnostics, pageRules);

    private static void ScanRawRules(string css, int start, int end, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, HtmlCssPageRuleSet pageRules) {
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
                    ScanRawRules(css, boundary + 1, closeBrace, options, diagnostics, pageRules);
                }
            } else if (string.Equals(name, "page", StringComparison.OrdinalIgnoreCase)) {
                string body = css.Substring(boundary + 1, closeBrace - boundary - 1);
                ApplyRawPageRule(prelude, body, options, diagnostics, pageRules);
            }

            cursor = closeBrace + 1;
        }
    }

    private static void ApplyRawPageRule(string selectorText, string body, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, HtmlCssPageRuleSet pageRules) {
        if (!TryParsePageSelector(selectorText, out HtmlCssPageSelector selector)) {
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSelectorPending, "A named or compound page selector is not yet applied to individual pages.", HtmlDiagnosticSeverity.Warning, selectorText.Length == 0 ? "@page" : selectorText);
            return;
        }

        string size = FindTopLevelDeclaration(body, "size");
        if (selector == HtmlCssPageSelector.Generic) {
            if (size.Length > 0 && !TryApplyPageSize(size, options)) {
                diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSizeUnsupported, "The @page size declaration could not be mapped to a supported physical page size.", HtmlDiagnosticSeverity.Warning, "@page", size);
            }
        } else if (size.Length > 0 || HasPageMarginDeclaration(body)) {
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PagePseudoGeometryPending, "Pseudo-page size and margin declarations require page-by-page body reflow and were not applied to body geometry.", HtmlDiagnosticSeverity.Warning, "@page " + selectorText);
        }

        IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> marginBoxes = ExtractMarginBoxes(body, selectorText, options, diagnostics);
        if (marginBoxes.Count > 0) pageRules.Add(new HtmlCssPageRule(selector, marginBoxes));
    }

    private static bool TryParsePageSelector(string selectorText, out HtmlCssPageSelector selector) {
        string normalized = selectorText.Trim();
        if (normalized.Length == 0) {
            selector = HtmlCssPageSelector.Generic;
            return true;
        }

        if (string.Equals(normalized, ":first", StringComparison.OrdinalIgnoreCase)) selector = HtmlCssPageSelector.First;
        else if (string.Equals(normalized, ":left", StringComparison.OrdinalIgnoreCase)) selector = HtmlCssPageSelector.Left;
        else if (string.Equals(normalized, ":right", StringComparison.OrdinalIgnoreCase)) selector = HtmlCssPageSelector.Right;
        else {
            selector = HtmlCssPageSelector.Generic;
            return false;
        }

        return true;
    }

    private static bool HasPageMarginDeclaration(string body) =>
        FindTopLevelDeclaration(body, "margin").Length > 0
        || FindTopLevelDeclaration(body, "margin-top").Length > 0
        || FindTopLevelDeclaration(body, "margin-right").Length > 0
        || FindTopLevelDeclaration(body, "margin-bottom").Length > 0
        || FindTopLevelDeclaration(body, "margin-left").Length > 0;

    private static IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> ExtractMarginBoxes(string pageBody, string pageSelector, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        var boxes = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate>();
        int cursor = 0;
        while (cursor < pageBody.Length) {
            if (IsCommentStart(pageBody, cursor)) {
                cursor = SkipComment(pageBody, cursor + 2, pageBody.Length);
                continue;
            }

            char current = pageBody[cursor];
            if (current == '\'' || current == '"') {
                cursor = SkipQuoted(pageBody, cursor + 1, pageBody.Length, current);
                continue;
            }

            if (current != '@') {
                cursor++;
                continue;
            }

            int nameStart = cursor + 1;
            int nameEnd = nameStart;
            while (nameEnd < pageBody.Length && (char.IsLetter(pageBody[nameEnd]) || pageBody[nameEnd] == '-')) nameEnd++;
            string name = pageBody.Substring(nameStart, nameEnd - nameStart).ToLowerInvariant();
            int boundary = FindRuleBoundary(pageBody, nameEnd, pageBody.Length);
            if (boundary < 0 || pageBody[boundary] == ';') {
                cursor = boundary < 0 ? pageBody.Length : boundary + 1;
                continue;
            }

            int close = FindMatchingBrace(pageBody, boundary);
            if (close < 0) break;
            string marginBody = pageBody.Substring(boundary + 1, close - boundary - 1);
            if (!TryMapMarginPosition(name, out HtmlCssPageMarginPosition position)) {
                diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageMarginPositionUnsupported, "A page-margin position is not recognized by the direct renderer.", HtmlDiagnosticSeverity.Warning, "@page " + pageSelector, "@" + name);
                cursor = close + 1;
                continue;
            }

            string contentValue = FindTopLevelDeclaration(marginBody, "content");
            if (!HtmlCssGeneratedContentTemplate.TryParse(contentValue, out HtmlCssGeneratedContentTemplate content)) {
                diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageMarginContentUnsupported, "A page-margin content expression could not be represented.", HtmlDiagnosticSeverity.Warning, "@page " + pageSelector + " @" + name, contentValue);
                cursor = close + 1;
                continue;
            }

            boxes[position] = CreateMarginTemplate(position, content, marginBody, options);
            cursor = close + 1;
        }

        return boxes;
    }

    private static HtmlCssPageMarginTemplate CreateMarginTemplate(HtmlCssPageMarginPosition position, HtmlCssGeneratedContentTemplate content, string body, HtmlRenderOptions options) {
        string family = HtmlRenderCssValues.FirstFontFamily(FindTopLevelDeclaration(body, "font-family"), options.DefaultFontFamily);
        double fontSize = options.DefaultFontSize;
        HtmlRenderCssValues.TryLength(FindTopLevelDeclaration(body, "font-size"), options.DefaultFontSize, options.DefaultFontSize, options.DefaultFontSize, out fontSize);
        if (fontSize <= 0D) fontSize = options.DefaultFontSize;
        OfficeFontStyle fontStyle = OfficeFontStyle.Regular;
        string weight = FindTopLevelDeclaration(body, "font-weight");
        if (string.Equals(weight, "bold", StringComparison.OrdinalIgnoreCase) || int.TryParse(weight, out int numericWeight) && numericWeight >= 600) fontStyle |= OfficeFontStyle.Bold;
        string style = FindTopLevelDeclaration(body, "font-style");
        if (style.StartsWith("italic", StringComparison.OrdinalIgnoreCase) || style.StartsWith("oblique", StringComparison.OrdinalIgnoreCase)) fontStyle |= OfficeFontStyle.Italic;
        OfficeColor color = HtmlRenderCssValues.TryColor(FindTopLevelDeclaration(body, "color"), out OfficeColor parsedColor) ? parsedColor : OfficeColor.Black;
        OfficeTextAlignment alignment = ResolveMarginAlignment(position, FindTopLevelDeclaration(body, "text-align"));
        return new HtmlCssPageMarginTemplate(position, content, new OfficeFontInfo(family, fontSize, fontStyle), color, alignment);
    }

    private static OfficeTextAlignment ResolveMarginAlignment(HtmlCssPageMarginPosition position, string value) {
        if (string.Equals(value, "left", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Left;
        if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Center;
        if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Right;
        if (position == HtmlCssPageMarginPosition.TopCenter || position == HtmlCssPageMarginPosition.BottomCenter) return OfficeTextAlignment.Center;
        if (position == HtmlCssPageMarginPosition.LeftTop || position == HtmlCssPageMarginPosition.LeftMiddle || position == HtmlCssPageMarginPosition.LeftBottom
            || position == HtmlCssPageMarginPosition.RightTop || position == HtmlCssPageMarginPosition.RightMiddle || position == HtmlCssPageMarginPosition.RightBottom) return OfficeTextAlignment.Center;
        if (position == HtmlCssPageMarginPosition.TopRight || position == HtmlCssPageMarginPosition.TopRightCorner
            || position == HtmlCssPageMarginPosition.BottomRight || position == HtmlCssPageMarginPosition.BottomRightCorner) return OfficeTextAlignment.Right;
        return OfficeTextAlignment.Left;
    }

    private static bool TryMapMarginPosition(string name, out HtmlCssPageMarginPosition position) {
        switch (name) {
            case "top-left-corner": position = HtmlCssPageMarginPosition.TopLeftCorner; return true;
            case "top-left": position = HtmlCssPageMarginPosition.TopLeft; return true;
            case "top-center": position = HtmlCssPageMarginPosition.TopCenter; return true;
            case "top-right": position = HtmlCssPageMarginPosition.TopRight; return true;
            case "top-right-corner": position = HtmlCssPageMarginPosition.TopRightCorner; return true;
            case "left-top": position = HtmlCssPageMarginPosition.LeftTop; return true;
            case "left-middle": position = HtmlCssPageMarginPosition.LeftMiddle; return true;
            case "left-bottom": position = HtmlCssPageMarginPosition.LeftBottom; return true;
            case "right-top": position = HtmlCssPageMarginPosition.RightTop; return true;
            case "right-middle": position = HtmlCssPageMarginPosition.RightMiddle; return true;
            case "right-bottom": position = HtmlCssPageMarginPosition.RightBottom; return true;
            case "bottom-left-corner": position = HtmlCssPageMarginPosition.BottomLeftCorner; return true;
            case "bottom-left": position = HtmlCssPageMarginPosition.BottomLeft; return true;
            case "bottom-center": position = HtmlCssPageMarginPosition.BottomCenter; return true;
            case "bottom-right": position = HtmlCssPageMarginPosition.BottomRight; return true;
            case "bottom-right-corner": position = HtmlCssPageMarginPosition.BottomRightCorner; return true;
            default: position = HtmlCssPageMarginPosition.TopLeft; return false;
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
