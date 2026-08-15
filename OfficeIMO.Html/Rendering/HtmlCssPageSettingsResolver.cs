using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssPageSettingsResolver {
    internal static HtmlCssPageRuleSet Apply(IHtmlDocument document, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        var pageRules = new HtmlCssPageRuleSet(options);
        if (options.Mode != HtmlRenderMode.Paged || !options.HonorCssPageRules) return pageRules;
        var layers = new CascadeLayerRegistry();
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !IsApplicablePrintMedia(styleElement.GetAttribute("media") ?? string.Empty, options)) continue;
            ApplyRawPageRules(styleElement.TextContent, options, diagnostics, pageRules, layers);
        }
        pageRules.ApplyGenericGeometry(options);
        return pageRules;
    }

    internal static bool TryResolvePageSize(string value, double currentWidth, double currentHeight, double fontSize, out double width, out double height) {
        width = currentWidth;
        height = currentHeight;
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        if (parts.Count == 0 || parts.Count > 2) return false;
        int landscapeCount = parts.Count(part => string.Equals(part, "landscape", StringComparison.OrdinalIgnoreCase));
        int portraitCount = parts.Count(part => string.Equals(part, "portrait", StringComparison.OrdinalIgnoreCase));
        bool landscape = landscapeCount == 1;
        bool portrait = portraitCount == 1;
        bool automatic = parts.Any(part => string.Equals(part, "auto", StringComparison.OrdinalIgnoreCase));
        if (automatic) {
            if (parts.Count != 1 || landscape || portrait) return false;
            return true;
        }
        if (landscapeCount > 1 || portraitCount > 1 || landscape && portrait) return false;

        OfficePageSize? named = null;
        int namedCount = 0;
        foreach (string part in parts) {
            OfficePageSize? candidate = ResolveNamedSize(part);
            if (!candidate.HasValue) continue;
            named = candidate;
            namedCount++;
        }
        int orientationCount = landscapeCount + portraitCount;
        if (namedCount > 0) {
            if (namedCount != 1 || parts.Count != namedCount + orientationCount) return false;
            OfficePageSize namedSize = named.GetValueOrDefault();
            OfficePageSize resolved = landscape ? namedSize.Landscape() : portrait ? namedSize.Portrait() : namedSize;
            width = resolved.WidthInches * HtmlRenderOptions.CssPixelsPerInch;
            height = resolved.HeightInches * HtmlRenderOptions.CssPixelsPerInch;
            return true;
        }

        if (orientationCount > 0) {
            if (parts.Count != 1) return false;
            width = landscape ? Math.Max(currentWidth, currentHeight) : Math.Min(currentWidth, currentHeight);
            height = landscape ? Math.Min(currentWidth, currentHeight) : Math.Max(currentWidth, currentHeight);
            return true;
        }

        var lengths = new List<double>();
        foreach (string part in parts) {
            if (!HtmlRenderCssValues.HasExplicitLengthSyntax(part, allowPercentage: false, allowUnitlessZero: false)
                || !HtmlRenderCssValues.TryLength(part, currentWidth, fontSize, fontSize, currentWidth, currentHeight, out double length)
                || length <= 0D) {
                return false;
            }
            lengths.Add(length);
        }

        if (lengths.Count != 2) return false;
        var custom = new OfficePageSize(lengths[0] / HtmlRenderOptions.CssPixelsPerInch, lengths[1] / HtmlRenderOptions.CssPixelsPerInch);
        OfficePageSize customResolved = landscape ? custom.Landscape() : portrait ? custom.Portrait() : custom;
        width = customResolved.WidthInches * HtmlRenderOptions.CssPixelsPerInch;
        height = customResolved.HeightInches * HtmlRenderOptions.CssPixelsPerInch;
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

    private static bool IsCssStyleElement(IElement element) {
        return HtmlResourcePipeline.IsCssStyleElement(element);
    }

    private static bool IsApplicablePrintMedia(string mediaText, HtmlRenderOptions options) =>
        HtmlComputedStyleEngine.IsApplicableMedia(
            mediaText,
            HtmlCssMediaContext.Print,
            options.PageWidth,
            options.PageHeight,
            options.MediaFeatures);

    private static void ApplyRawPageRules(
        string css,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        HtmlCssPageRuleSet pageRules,
        CascadeLayerRegistry layers) =>
        ScanRawRules(css, 0, css.Length, options, diagnostics, pageRules, layers, null, null);

    private static void ScanRawRules(
        string css,
        int start,
        int end,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        HtmlCssPageRuleSet pageRules,
        CascadeLayerRegistry layers,
        string? layerPath,
        CascadeLayerOrder? layerOrder) {
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
                if (boundary >= 0 && string.Equals(name, "layer", StringComparison.OrdinalIgnoreCase)) {
                    layers.RegisterStatement(css.Substring(nameEnd, boundary - nameEnd), layerPath);
                }
                cursor = boundary < 0 ? end : boundary + 1;
                continue;
            }

            int closeBrace = FindMatchingBrace(css, boundary);
            if (closeBrace < 0 || closeBrace >= end) return;
            string prelude = css.Substring(nameEnd, boundary - nameEnd).Trim();
            if (string.Equals(name, "media", StringComparison.OrdinalIgnoreCase)) {
                if (IsApplicablePrintMedia(prelude, options)) {
                    ScanRawRules(css, boundary + 1, closeBrace, options, diagnostics, pageRules, layers, layerPath, layerOrder);
                }
            } else if (string.Equals(name, "supports", StringComparison.OrdinalIgnoreCase)) {
                if (HtmlComputedStyleEngine.IsApplicableSupports(prelude)) {
                    ScanRawRules(css, boundary + 1, closeBrace, options, diagnostics, pageRules, layers, layerPath, layerOrder);
                }
            } else if (string.Equals(name, "layer", StringComparison.OrdinalIgnoreCase)) {
                (string nestedPath, CascadeLayerOrder nestedOrder) = layers.RegisterBlock(prelude, layerPath);
                ScanRawRules(css, boundary + 1, closeBrace, options, diagnostics, pageRules, layers, nestedPath, nestedOrder);
            } else if (string.Equals(name, "page", StringComparison.OrdinalIgnoreCase)) {
                string body = css.Substring(boundary + 1, closeBrace - boundary - 1);
                ApplyRawPageRule(prelude, body, options, diagnostics, pageRules, layerOrder);
            }

            cursor = closeBrace + 1;
        }
    }

    private static void ApplyRawPageRule(string selectorText, string body, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, HtmlCssPageRuleSet pageRules, CascadeLayerOrder? layerOrder) {
        if (!TryParsePageSelector(selectorText, out string? pageName, out HtmlCssPageSelector selector)) {
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSelectorPending, "A complex page selector could not be applied to individual pages.", HtmlDiagnosticSeverity.Warning, selectorText.Length == 0 ? "@page" : selectorText);
            return;
        }

        bool IsValidSize(string value) => string.Equals(value.Trim(), "revert-layer", StringComparison.OrdinalIgnoreCase)
            || TryResolvePageSize(
                value,
                options.PageWidth,
                options.PageHeight,
                options.DefaultFontSize,
                out _,
                out _);
        HtmlCssPageDeclaration authoredSize = FindTopLevelDeclarationWithPriority(body, "size");
        HtmlCssPageDeclaration sizeDeclaration = FindTopLevelDeclarationWithPriority(body, "size", IsValidSize);
        if (authoredSize.Value.Length > 0 && !IsValidSize(authoredSize.Value)) {
            string source = selectorText.Length == 0 ? "@page" : "@page " + selectorText;
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSizeUnsupported, "The @page size declaration could not be mapped to a supported physical page size.", HtmlDiagnosticSeverity.Warning, source, authoredSize.Value);
        }
        var geometry = new HtmlCssPageGeometryDeclaration(
            sizeDeclaration,
            FindTopLevelDeclarationWithPriority(body, "margin", value => HtmlCssPageRuleSet.TryExpandMargin(value, out _)),
            FindTopLevelDeclarationWithPriority(body, "margin-top", HtmlCssPageRuleSet.IsValidPageMarginComponent),
            FindTopLevelDeclarationWithPriority(body, "margin-right", HtmlCssPageRuleSet.IsValidPageMarginComponent),
            FindTopLevelDeclarationWithPriority(body, "margin-bottom", HtmlCssPageRuleSet.IsValidPageMarginComponent),
            FindTopLevelDeclarationWithPriority(body, "margin-left", HtmlCssPageRuleSet.IsValidPageMarginComponent));
        IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> marginBoxes = ExtractMarginBoxes(body, selectorText, diagnostics);
        if (marginBoxes.Count > 0 || !geometry.IsEmpty) pageRules.Add(new HtmlCssPageRule(pageName, selector, marginBoxes, geometry, layerOrder));
    }

    private static bool TryParsePageSelector(string selectorText, out string? pageName, out HtmlCssPageSelector selector) {
        string normalized = selectorText.Trim();
        if (normalized.Length == 0) {
            pageName = null;
            selector = HtmlCssPageSelector.Generic;
            return true;
        }

        int pseudoStart = normalized.IndexOf(':');
        string name = pseudoStart < 0 ? normalized : normalized.Substring(0, pseudoStart).Trim();
        string pseudo = pseudoStart < 0 ? string.Empty : normalized.Substring(pseudoStart).Trim();
        if (name.Length > 0 && !IsPageIdentifier(name)) {
            pageName = null;
            selector = HtmlCssPageSelector.Generic;
            return false;
        }

        pageName = name.Length == 0 ? null : name;
        if (pseudo.Length == 0) selector = HtmlCssPageSelector.Generic;
        else if (string.Equals(pseudo, ":first", StringComparison.OrdinalIgnoreCase)) selector = HtmlCssPageSelector.First;
        else if (string.Equals(pseudo, ":left", StringComparison.OrdinalIgnoreCase)) selector = HtmlCssPageSelector.Left;
        else if (string.Equals(pseudo, ":right", StringComparison.OrdinalIgnoreCase)) selector = HtmlCssPageSelector.Right;
        else {
            selector = HtmlCssPageSelector.Generic;
            return false;
        }
        return true;
    }

    private static bool IsPageIdentifier(string value) {
        if (value.Length == 0 || !(char.IsLetter(value[0]) || value[0] == '_' || value[0] == '-')) return false;
        for (int index = 1; index < value.Length; index++) {
            char current = value[index];
            if (!(char.IsLetterOrDigit(current) || current == '_' || current == '-')) return false;
        }

        return !string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasPageMarginDeclaration(string body) =>
        FindTopLevelDeclaration(body, "margin").Length > 0
        || FindTopLevelDeclaration(body, "margin-top").Length > 0
        || FindTopLevelDeclaration(body, "margin-right").Length > 0
        || FindTopLevelDeclaration(body, "margin-bottom").Length > 0
        || FindTopLevelDeclaration(body, "margin-left").Length > 0;

    private static IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> ExtractMarginBoxes(string pageBody, string pageSelector, HtmlDiagnosticReport diagnostics) {
        var boxes = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule>();
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

            HtmlCssPageDeclaration authoredContent = FindTopLevelDeclarationWithPriority(marginBody, "content");
            HtmlCssPageDeclaration content = FindTopLevelDeclarationWithPriority(
                marginBody,
                "content",
                value => IsRevertLayer(value) || HtmlCssGeneratedContentTemplate.TryParse(value, out _));
            if (authoredContent.Value.Length > 0
                && !IsRevertLayer(authoredContent.Value)
                && !HtmlCssGeneratedContentTemplate.TryParse(authoredContent.Value, out _)) {
                diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageMarginContentUnsupported, "A page-margin content expression could not be represented.", HtmlDiagnosticSeverity.Warning, "@page " + pageSelector + " @" + name, authoredContent.Value);
            }

            ICssStyleDeclaration? style = new CssParser().ParseDeclaration(marginBody);
            var marginRule = new HtmlCssPageMarginRule(
                content,
                ReadStyleDeclaration(style, marginBody, "font-family", cursor),
                ReadStyleDeclaration(style, marginBody, "font-size", cursor),
                ReadStyleDeclaration(style, marginBody, "font-weight", cursor),
                ReadStyleDeclaration(style, marginBody, "font-style", cursor),
                ReadStyleDeclaration(style, marginBody, "color", cursor),
                ReadStyleDeclaration(style, marginBody, "text-align", cursor));
            if (!marginRule.IsEmpty) {
                boxes[position] = boxes.TryGetValue(position, out HtmlCssPageMarginRule? earlier)
                    ? HtmlCssPageMarginRule.Merge(earlier, marginRule)
                    : marginRule;
            }
            cursor = close + 1;
        }

        return boxes;
    }

    private static bool IsRevertLayer(string value) =>
        string.Equals(value.Trim(), "revert-layer", StringComparison.OrdinalIgnoreCase);

    private static HtmlCssPageDeclaration ReadStyleDeclaration(ICssStyleDeclaration? style, string body, string propertyName, int order) {
        HtmlCssPageDeclaration authored = FindTopLevelDeclarationWithPriority(body, propertyName);
        if (IsRevertLayer(authored.Value)) return authored;
        if (style == null) return new HtmlCssPageDeclaration(string.Empty, false, order);
        string value = style.GetPropertyValue(propertyName);
        bool important = string.Equals(style.GetPropertyPriority(propertyName), "important", StringComparison.OrdinalIgnoreCase);
        return new HtmlCssPageDeclaration(value, important, order);
    }

    internal static OfficeTextAlignment ResolveMarginAlignment(HtmlCssPageMarginPosition position, string value) {
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

    private static string FindTopLevelDeclaration(string body, string propertyName) =>
        FindTopLevelDeclarationWithPriority(body, propertyName).Value;

    private static HtmlCssPageDeclaration FindTopLevelDeclarationWithPriority(
        string body,
        string propertyName,
        Func<string, bool>? isValid = null) {
        body = HtmlComputedStyleEngine.StripCssCommentsOutsideStrings(body);
        int start = 0;
        int depth = 0;
        char quote = '\0';
        var resolved = new HtmlCssPageDeclaration(string.Empty, false, -1);
        int declarationOrder = 0;
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
                    string value = declaration.Substring(separator + 1).Trim();
                    bool important = TryStripImportant(ref value);
                    if ((isValid == null || isValid(value)) && (important || !resolved.IsImportant)) {
                        resolved = new HtmlCssPageDeclaration(value, important, declarationOrder);
                    }
                }

                start = index + 1;
                declarationOrder++;
            }
        }

        return resolved;
    }

    private static bool TryStripImportant(ref string value) {
        int end = value.Length;
        while (end > 0 && char.IsWhiteSpace(value[end - 1])) end--;
        const string Important = "important";
        int wordStart = end - Important.Length;
        if (wordStart < 0 || !string.Equals(value.Substring(wordStart, Important.Length), Important, StringComparison.OrdinalIgnoreCase)) return false;
        int bang = wordStart - 1;
        while (bang >= 0 && char.IsWhiteSpace(value[bang])) bang--;
        if (bang < 0 || value[bang] != '!') return false;
        value = value.Substring(0, bang).TrimEnd();
        return true;
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
