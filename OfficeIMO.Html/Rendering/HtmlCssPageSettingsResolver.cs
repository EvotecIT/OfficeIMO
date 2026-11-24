using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

internal static partial class HtmlCssPageSettingsResolver {
    internal static HtmlCssPageRuleSet Apply(IHtmlDocument document, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics,
        HtmlConversionLimits? limits = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var fontBudget = new HtmlCssProcessingBudget(limits);
        var pageRules = new HtmlCssPageRuleSet(options);
        if (options.Mode != HtmlRenderMode.Paged || !options.HonorCssPageRules) return pageRules;
        var layers = new CascadeLayerRegistry();
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !IsApplicablePageMedia(styleElement.GetAttribute("media") ?? string.Empty, options)) continue;
            string css = HtmlStyleElementText.Read(styleElement);
            SyntaxCache? cache = styleElement.HasAttribute("data-officeimo-source") ? options.PublicationPageRules : null;
            IEnumerable<RawPageRule> rules = cache?.Get(options.ExternalStylesheets?.ReuseMaterializedText(css) ?? css, cancellationToken)
                ?? ReadRawPageRules(css, 0, css.Length, cancellationToken);
            ApplyRawPageRules(rules, options, diagnostics, pageRules, layers, null, null, fontBudget, cache != null, cancellationToken);
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
            width = resolved.Width.ToInches() * HtmlRenderOptions.CssPixelsPerInch;
            height = resolved.Height.ToInches() * HtmlRenderOptions.CssPixelsPerInch;
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
        var custom = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(lengths[0] / HtmlRenderOptions.CssPixelsPerInch), global::OfficeIMO.OfficeLength.Inches(lengths[1] / HtmlRenderOptions.CssPixelsPerInch));
        OfficePageSize customResolved = landscape ? custom.Landscape() : portrait ? custom.Portrait() : custom;
        width = customResolved.Width.ToInches() * HtmlRenderOptions.CssPixelsPerInch;
        height = customResolved.Height.ToInches() * HtmlRenderOptions.CssPixelsPerInch;
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

    private static bool IsApplicablePageMedia(string mediaText, HtmlRenderOptions options) =>
        HtmlComputedStyleEngine.IsApplicableMedia(
            mediaText,
            options.MediaContext,
            options.PageWidth,
            options.PageHeight,
            options.MediaFeatures);

    private static void ApplyRawPageRule(RawPageRule rule, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, HtmlCssPageRuleSet pageRules, CascadeLayerOrder? layerOrder,
        HtmlCssProcessingBudget fontBudget, CancellationToken cancellationToken) {
        string selectorText = rule.Prelude;
        if (!TryParsePageSelector(selectorText, out string? pageName, out HtmlCssPageSelector selector)) {
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSelectorPending, "A complex page selector could not be applied to individual pages.", HtmlDiagnosticSeverity.Warning, selectorText.Length == 0 ? "@page" : selectorText);
            return;
        }
        PageRuleBody body = rule.ReadPageBody(fontBudget, cancellationToken);
        PageDeclarations declarations = body.Declarations;

        bool IsValidSize(string value) => string.Equals(value.Trim(), "revert-layer", StringComparison.OrdinalIgnoreCase)
            || TryResolvePageSize(
                value,
                options.PageWidth,
                options.PageHeight,
                options.DefaultFontSize,
                out _,
                out _);
        HtmlCssPageDeclaration authoredSize = declarations.Find("size");
        HtmlCssPageDeclaration sizeDeclaration = declarations.Find("size", IsValidSize);
        if (authoredSize.Value.Length > 0 && !IsValidSize(authoredSize.Value)) {
            string source = selectorText.Length == 0 ? "@page" : "@page " + selectorText;
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageSizeUnsupported, "The @page size declaration could not be mapped to a supported physical page size.", HtmlDiagnosticSeverity.Warning, source, authoredSize.Value);
        }
        HtmlCssPageDeclaration backgroundColor = ResolvePageBackgroundDeclaration(declarations);
        var geometry = new HtmlCssPageGeometryDeclaration(
            sizeDeclaration,
            declarations.Find("margin", value => HtmlCssPageRuleSet.TryExpandMargin(value, out _)),
            declarations.Find("margin-top", HtmlCssPageRuleSet.IsValidPageMarginComponent),
            declarations.Find("margin-right", HtmlCssPageRuleSet.IsValidPageMarginComponent),
            declarations.Find("margin-bottom", HtmlCssPageRuleSet.IsValidPageMarginComponent),
            declarations.Find("margin-left", HtmlCssPageRuleSet.IsValidPageMarginComponent),
            backgroundColor);
        bool IsValidBleed(string value) => IsCssWidePageProductionValue(value)
            || string.Equals(value.Trim(), "auto", StringComparison.OrdinalIgnoreCase)
            || HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: false, allowUnitlessZero: true)
                && HtmlRenderCssValues.TryLength(value, 1D, options.DefaultFontSize, options.DefaultFontSize, options.PageWidth, options.PageHeight, out double bleed)
                && bleed >= 0D;
        bool IsValidMarks(string value) {
            if (IsCssWidePageProductionValue(value)) return true;
            IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
            if (tokens.Count == 0 || tokens.Count > 2) return false;
            if (tokens.Count == 1 && string.Equals(tokens[0], "none", StringComparison.OrdinalIgnoreCase)) return true;
            return tokens.All(token => string.Equals(token, "crop", StringComparison.OrdinalIgnoreCase)
                || string.Equals(token, "cross", StringComparison.OrdinalIgnoreCase))
                && tokens.Distinct(StringComparer.OrdinalIgnoreCase).Count() == tokens.Count;
        }
        HtmlCssPageDeclaration authoredBleed = declarations.Find("bleed");
        HtmlCssPageDeclaration authoredMarks = declarations.Find("marks");
        string productionSource = selectorText.Length == 0 ? "@page" : "@page " + selectorText;
        if (authoredBleed.Value.Length > 0 && !IsValidBleed(authoredBleed.Value)) {
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageBleedUnsupported,
                "An @page bleed declaration could not be resolved.", HtmlDiagnosticSeverity.Warning,
                productionSource, authoredBleed.Value, OfficeConversionLossKind.Omission);
        }
        if (authoredMarks.Value.Length > 0 && !IsValidMarks(authoredMarks.Value)) {
            diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageMarksUnsupported,
                "An @page marks declaration could not be resolved.", HtmlDiagnosticSeverity.Warning,
                productionSource, authoredMarks.Value, OfficeConversionLossKind.Omission);
        }
        var production = new HtmlCssPageProductionDeclaration(
            declarations.Find("bleed", IsValidBleed),
            declarations.Find("marks", IsValidMarks));
        diagnostics.AddRange(body.Diagnostics);
        if (body.MarginBoxes.Count > 0 || body.StyleDeclarations.Count > 0 || !geometry.IsEmpty || !production.IsEmpty) {
            pageRules.Add(new HtmlCssPageRule(pageName, selector, body.MarginBoxes, geometry, production, layerOrder, body.StyleDeclarations));
        }
    }

    private static HtmlCssPageDeclaration ResolvePageBackgroundDeclaration(PageDeclarations declarations) {
        static bool IsSupported(string value) => IsCssWidePageProductionValue(value)
            || string.Equals(value.Trim(), "transparent", StringComparison.OrdinalIgnoreCase)
            || HtmlRenderCssValues.TryColor(value.Trim(), out _);

        HtmlCssPageDeclaration shorthand = declarations.Find("background", IsSupported);
        HtmlCssPageDeclaration longhand = declarations.Find("background-color", IsSupported);
        if (shorthand.Value.Length == 0) return longhand;
        if (longhand.Value.Length == 0) return shorthand;
        if (shorthand.IsImportant != longhand.IsImportant) return longhand.IsImportant ? longhand : shorthand;
        return longhand.Order >= shorthand.Order ? longhand : shorthand;
    }

    private static bool IsCssWidePageProductionValue(string value) =>
        string.Equals(value.Trim(), "initial", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value.Trim(), "unset", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value.Trim(), "revert-layer", StringComparison.OrdinalIgnoreCase);

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

    private static IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> ExtractMarginBoxes(string pageBody, string pageSelector, HtmlDiagnosticReport diagnostics,
        HtmlCssProcessingBudget fontBudget, CancellationToken cancellationToken) {
        var boxes = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule>();
        int cursor = 0;
        while (cursor < pageBody.Length) {
            cancellationToken.ThrowIfCancellationRequested();
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
            int boundary = FindRuleBoundary(pageBody, nameEnd, pageBody.Length, cancellationToken);
            if (boundary < 0 || pageBody[boundary] == ';') {
                cursor = boundary < 0 ? pageBody.Length : boundary + 1;
                continue;
            }

            int close = FindMatchingBrace(pageBody, boundary, cancellationToken);
            if (close < 0) break;
            string marginBody = pageBody.Substring(boundary + 1, close - boundary - 1);
            if (!TryMapMarginPosition(name, out HtmlCssPageMarginPosition position)) {
                diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageMarginPositionUnsupported, "A page-margin position is not recognized by the direct renderer.", HtmlDiagnosticSeverity.Warning, "@page " + pageSelector, "@" + name);
                cursor = close + 1;
                continue;
            }

            HtmlCssStyleBlock syntax = ParseContextStyleBlock(marginBody, fontBudget, cancellationToken);
            var declarations = new PageDeclarations(syntax, isMarginBox: true, cancellationToken);
            HtmlCssPageDeclaration authoredContent = declarations.Find("content");
            HtmlCssPageDeclaration content = declarations.Find(
                "content",
                value => IsRevertLayer(value) || HtmlCssGeneratedContentTemplate.TryParse(value, out _));
            if (authoredContent.Value.Length > 0
                && !IsRevertLayer(authoredContent.Value)
                && !HtmlCssGeneratedContentTemplate.TryParse(authoredContent.Value, out _)) {
                diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.PageMarginContentUnsupported, "A page-margin content expression could not be represented.", HtmlDiagnosticSeverity.Warning, "@page " + pageSelector + " @" + name, authoredContent.Value);
            }

            var marginRule = new HtmlCssPageMarginRule(content,
                ReadContextStyleDeclarations(syntax, cancellationToken));
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



    internal static OfficeTextAlignment ResolveMarginAlignment(HtmlCssPageMarginPosition position, string value) {
        if (string.Equals(value, "left", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Left;
        if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Center;
        if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Right;
        if (position == HtmlCssPageMarginPosition.TopCenter || position == HtmlCssPageMarginPosition.BottomCenter) return OfficeTextAlignment.Center;
        if (position == HtmlCssPageMarginPosition.LeftTop || position == HtmlCssPageMarginPosition.LeftMiddle || position == HtmlCssPageMarginPosition.LeftBottom
            || position == HtmlCssPageMarginPosition.RightTop || position == HtmlCssPageMarginPosition.RightMiddle || position == HtmlCssPageMarginPosition.RightBottom) return OfficeTextAlignment.Center;
        if (position == HtmlCssPageMarginPosition.TopRight || position == HtmlCssPageMarginPosition.TopLeftCorner
            || position == HtmlCssPageMarginPosition.BottomRight || position == HtmlCssPageMarginPosition.BottomLeftCorner) return OfficeTextAlignment.Right;
        return OfficeTextAlignment.Left;
    }

    /// <summary>Resolves physical table-cell alignment, retaining unsupported values for the layout diagnostic owner.</summary>
    internal static OfficeTextVerticalAlignment ResolveMarginVerticalAlignment(HtmlCssPageMarginPosition position,
        string value, out string unsupportedValue) {
        unsupportedValue = string.Empty;
        switch (value.Trim().ToLowerInvariant()) {
            // Each margin box is an independent cell with no peer baseline to displace it.
            case "baseline":
            case "top": return OfficeTextVerticalAlignment.Top;
            case "middle": return OfficeTextVerticalAlignment.Center;
            case "bottom": return OfficeTextVerticalAlignment.Bottom;
            case "": break;
            default: unsupportedValue = value; break;
        }
        return position switch {
            HtmlCssPageMarginPosition.LeftTop or HtmlCssPageMarginPosition.RightTop => OfficeTextVerticalAlignment.Top,
            HtmlCssPageMarginPosition.LeftBottom or HtmlCssPageMarginPosition.RightBottom => OfficeTextVerticalAlignment.Bottom,
            _ => OfficeTextVerticalAlignment.Center
        };
    }

    /// <summary>Returns the canonical CSS spelling shared by generated styles, layout sources and diagnostics.</summary>
    internal static string GetMarginBoxName(HtmlCssPageMarginPosition position) => position switch {
        HtmlCssPageMarginPosition.TopLeftCorner => "top-left-corner",
        HtmlCssPageMarginPosition.TopLeft => "top-left",
        HtmlCssPageMarginPosition.TopCenter => "top-center",
        HtmlCssPageMarginPosition.TopRight => "top-right",
        HtmlCssPageMarginPosition.TopRightCorner => "top-right-corner",
        HtmlCssPageMarginPosition.LeftTop => "left-top",
        HtmlCssPageMarginPosition.LeftMiddle => "left-middle",
        HtmlCssPageMarginPosition.LeftBottom => "left-bottom",
        HtmlCssPageMarginPosition.RightTop => "right-top",
        HtmlCssPageMarginPosition.RightMiddle => "right-middle",
        HtmlCssPageMarginPosition.RightBottom => "right-bottom",
        HtmlCssPageMarginPosition.BottomLeftCorner => "bottom-left-corner",
        HtmlCssPageMarginPosition.BottomLeft => "bottom-left",
        HtmlCssPageMarginPosition.BottomCenter => "bottom-center",
        HtmlCssPageMarginPosition.BottomRight => "bottom-right",
        HtmlCssPageMarginPosition.BottomRightCorner => "bottom-right-corner",
        _ => throw new ArgumentOutOfRangeException(nameof(position))
    };

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

    private static int FindRuleBoundary(string css, int start, int end, CancellationToken cancellationToken) {
        int parentheses = 0;
        for (int index = start; index < end; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
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

    private static int FindMatchingBrace(string css, int open, CancellationToken cancellationToken) {
        int depth = 0;
        char quote = '\0';
        for (int index = open; index < css.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
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
