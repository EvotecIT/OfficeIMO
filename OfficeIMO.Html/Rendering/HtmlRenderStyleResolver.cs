using System.Globalization;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HtmlComputedStyleSet _computedStyles;
    private readonly HtmlRenderOptions _options;

    internal HtmlRenderStyleResolver(HtmlComputedStyleSet computedStyles, HtmlRenderOptions options) {
        _computedStyles = computedStyles;
        _options = options;
    }

    internal HtmlRenderBoxStyle Resolve(IElement element, double containingWidth, HtmlRenderBoxStyle? parent = null) {
        HtmlComputedStyle computed = _computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? found)
            ? found
            : new HtmlComputedStyle(new Dictionary<string, string>());
        return ResolveCore(element, computed, containingWidth, parent, false, string.Empty);
    }

    internal bool TryResolvePseudo(
        IElement element,
        HtmlPseudoElementKind kind,
        double containingWidth,
        HtmlRenderBoxStyle parent,
        out HtmlRenderBoxStyle style) {
        if (!_computedStyles.TryGetPseudoStyle(element, kind, out HtmlComputedStyle computed)) {
            style = null!;
            return false;
        }

        string semanticRole = kind == HtmlPseudoElementKind.Before ? "generated-before" : "generated-after";
        style = ResolveCore(element, computed, containingWidth, parent, true, semanticRole);
        return true;
    }

    private HtmlRenderBoxStyle ResolveCore(
        IElement element,
        HtmlComputedStyle computed,
        double containingWidth,
        HtmlRenderBoxStyle? parent,
        bool pseudoElement,
        string pseudoSemanticRole) {
        string tag = element.TagName.ToLowerInvariant();
        double parentFontSize = parent?.Font.Size ?? _options.DefaultFontSize;
        string fontSizeValue = computed.GetValue("font-size");
        double fontSize = string.IsNullOrWhiteSpace(fontSizeValue)
            ? (pseudoElement ? parentFontSize : ResolveDefaultTagFontSize(tag, parentFontSize))
            : ResolveFontSize(fontSizeValue, parentFontSize);
        OfficeFontStyle fontStyle = ResolveFontStyle(pseudoElement ? string.Empty : tag, computed);
        string defaultFamily = !pseudoElement && (tag == "code" || tag == "pre" || tag == "kbd" || tag == "samp")
            ? "Consolas"
            : parent?.Font.FamilyName ?? _options.DefaultFontFamily;
        string family = HtmlRenderCssValues.FontFamilyList(computed.GetValue("font-family"), defaultFamily);
        string direction = ResolveDirection(computed.GetValue("direction"), parent?.Direction);

        var style = new HtmlRenderBoxStyle {
            Display = pseudoElement ? ResolvePseudoDisplay(computed.GetValue("display")) : ResolveDisplay(tag, computed.GetValue("display")),
            DisplayWasSpecified = !string.IsNullOrWhiteSpace(computed.GetValue("display")),
            PaintVisible = ResolvePaintVisibility(computed.GetValue("visibility"), parent),
            Font = new OfficeFontInfo(family, fontSize, fontStyle),
            Color = ResolveColor(computed.GetValue("color"), parent?.Color ?? OfficeColor.Black),
            Alignment = ResolveAlignment(computed.GetValue("text-align"), direction),
            LineHeight = ResolveLineHeight(computed.GetValue("line-height"), fontSize),
            SemanticRole = pseudoElement ? pseudoSemanticRole : ResolveSemanticRole(tag),
            PreserveWhitespace = IsPreformatted(pseudoElement ? string.Empty : tag, computed.GetValue("white-space")),
            ListStyleType = ResolveListStyleType(computed),
            TextTransform = string.IsNullOrWhiteSpace(computed.GetValue("text-transform")) ? parent?.TextTransform ?? "none" : computed.GetValue("text-transform").Trim().ToLowerInvariant(),
            Direction = direction,
            OverflowWrap = ResolveOverflowWrap(computed.GetValue("overflow-wrap"), parent?.OverflowWrap),
            WordBreak = ResolveWordBreak(computed.GetValue("word-break"), parent?.WordBreak),
            BorderBox = string.Equals(computed.GetValue("box-sizing"), "border-box", StringComparison.OrdinalIgnoreCase)
        };

        if (!pseudoElement) ApplyDefaultMargins(tag, fontSize, style);
        ApplyBoxValues(computed, containingWidth, fontSize, style);
        ApplyDimensions(element, computed, containingWidth, fontSize, parent, style, !pseudoElement);
        ApplyReplacedElementValues(computed, fontSize, style);
        ApplyPaint(computed, style);
        ApplyOverflow(computed, style);
        ApplyFloat(computed, style);
        ApplyPositioning(computed, style);
        ApplyFlex(computed, containingWidth, fontSize, style);
        ApplyColumns(computed, containingWidth, fontSize, style);
        ApplyGrid(computed, style);
        ApplyTable(computed, style);
        ApplyBreaks(computed, style);
        return style;
    }

    private static string ResolveOverflowWrap(string value, string? inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited ?? "normal";
        if (normalized == "normal" || normalized == "break-word" || normalized == "anywhere") return normalized;
        return "normal";
    }

    private static string ResolveWordBreak(string value, string? inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited ?? "normal";
        if (normalized == "normal" || normalized == "break-all" || normalized == "keep-all" || normalized == "break-word") return normalized;
        return "normal";
    }

    private static bool ResolvePaintVisibility(string value, HtmlRenderBoxStyle? parent) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return parent?.PaintVisible ?? true;
        return normalized != "hidden" && normalized != "collapse";
    }

    private void ApplyOverflow(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        string shorthand = computed.GetValue("overflow");
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(shorthand);
        if (values.Count == 1) {
            style.OverflowX = NormalizeOverflow(values[0], out style.UnsupportedOverflowX);
            style.OverflowY = NormalizeOverflow(values[0], out style.UnsupportedOverflowY);
        } else if (values.Count == 2) {
            style.OverflowX = NormalizeOverflow(values[0], out style.UnsupportedOverflowX);
            style.OverflowY = NormalizeOverflow(values[1], out style.UnsupportedOverflowY);
        } else if (values.Count > 2) {
            style.UnsupportedOverflowX = shorthand.Trim();
            style.UnsupportedOverflowY = shorthand.Trim();
        }

        string overflowX = computed.GetValue("overflow-x");
        if (!string.IsNullOrWhiteSpace(overflowX)) {
            style.OverflowX = NormalizeOverflow(overflowX, out style.UnsupportedOverflowX);
        }
        string overflowY = computed.GetValue("overflow-y");
        if (!string.IsNullOrWhiteSpace(overflowY)) {
            style.OverflowY = NormalizeOverflow(overflowY, out style.UnsupportedOverflowY);
        }
        string overflowClipMargin = computed.GetValue("overflow-clip-margin");
        if (!string.IsNullOrWhiteSpace(overflowClipMargin)
            && !HtmlCssOverflowClipMarginParser.TryParse(
                overflowClipMargin,
                style.Font.Size,
                _options.DefaultFontSize,
                out style.OverflowClipMarginBox,
                out style.OverflowClipMargin)) {
            style.UnsupportedOverflowClipMargin = overflowClipMargin.Trim().ToLowerInvariant();
        }

        if (style.OverflowX == "visible" && style.OverflowY != "visible" && style.OverflowY != "clip") style.OverflowX = "auto";
        if (style.OverflowY == "visible" && style.OverflowX != "visible" && style.OverflowX != "clip") style.OverflowY = "auto";
        if (style.OverflowX == "clip" && style.OverflowY != "visible" && style.OverflowY != "clip") style.OverflowX = "hidden";
        if (style.OverflowY == "clip" && style.OverflowX != "visible" && style.OverflowX != "clip") style.OverflowY = "hidden";
    }

    private static string NormalizeOverflow(string value, out string unsupported) {
        unsupported = string.Empty;
        string normalized = string.IsNullOrWhiteSpace(value) ? "visible" : value.Trim().ToLowerInvariant();
        if (normalized == "visible" || normalized == "hidden" || normalized == "clip" || normalized == "auto" || normalized == "scroll") return normalized;
        unsupported = normalized;
        return "visible";
    }

    private void ApplyTable(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        string captionSide = computed.GetValue("caption-side").Trim().ToLowerInvariant();
        if (captionSide.Length == 0 || captionSide == "top") {
            style.CaptionSide = "top";
        } else if (captionSide == "bottom") {
            style.CaptionSide = "bottom";
        } else {
            style.UnsupportedCaptionSide = captionSide;
            style.CaptionSide = "top";
        }

        string tableLayout = computed.GetValue("table-layout").Trim().ToLowerInvariant();
        if (tableLayout.Length == 0 || tableLayout == "auto") {
            style.TableLayout = "auto";
        } else if (tableLayout == "fixed") {
            style.TableLayout = "fixed";
        } else {
            style.UnsupportedTableLayout = tableLayout;
            style.TableLayout = "auto";
        }

        string borderCollapse = computed.GetValue("border-collapse").Trim().ToLowerInvariant();
        if (borderCollapse.Length == 0 || borderCollapse == "separate") {
            style.BorderCollapse = "separate";
        } else if (borderCollapse == "collapse") {
            style.BorderCollapse = "collapse";
        } else {
            style.UnsupportedBorderCollapse = borderCollapse;
            style.BorderCollapse = "separate";
        }

        string borderSpacing = computed.GetValue("border-spacing");
        if (!string.IsNullOrWhiteSpace(borderSpacing)
            && !HtmlCssTableParser.TryParseBorderSpacing(borderSpacing, style.Font.Size, _options.DefaultFontSize, out style.BorderSpacingX, out style.BorderSpacingY)) {
            style.UnsupportedBorderSpacing = borderSpacing.Trim().ToLowerInvariant();
        }
    }

    private static void ApplyFloat(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        style.FloatSide = NormalizeFloatSide(computed.GetValue("float"), style.Direction, out style.UnsupportedFloat);
        style.ClearSide = NormalizeClearSide(computed.GetValue("clear"), style.Direction, out style.UnsupportedClear);
    }

    private static string NormalizeFloatSide(string value, string direction, out string unsupported) {
        unsupported = string.Empty;
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value.Trim().ToLowerInvariant();
        if (normalized == "none" || normalized == "left" || normalized == "right") return normalized;
        if (normalized == "inline-start") return direction == "rtl" ? "right" : "left";
        if (normalized == "inline-end") return direction == "rtl" ? "left" : "right";
        unsupported = normalized;
        return "none";
    }

    private static string NormalizeClearSide(string value, string direction, out string unsupported) {
        unsupported = string.Empty;
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value.Trim().ToLowerInvariant();
        if (normalized == "none" || normalized == "left" || normalized == "right" || normalized == "both") return normalized;
        if (normalized == "inline-start") return direction == "rtl" ? "right" : "left";
        if (normalized == "inline-end") return direction == "rtl" ? "left" : "right";
        unsupported = normalized;
        return "none";
    }

    private static string ResolveDirection(string value, string? inherited) {
        string normalized = string.IsNullOrWhiteSpace(value) ? inherited ?? "ltr" : value.Trim().ToLowerInvariant();
        return normalized == "rtl" ? "rtl" : "ltr";
    }

    internal static bool IsBlockElement(IElement element, HtmlRenderBoxStyle style) {
        string display = style.Display;
        if (display == "none" || display == "contents") return false;
        if (display == "block" || display == "table" || display == "list-item" || display == "flex" || display == "grid" || display == "flow-root") return true;
        if (display == "inline" || display == "inline-block" || display == "inline-flex" || display == "inline-grid") return false;
        return IsDefaultBlockTag(element.TagName);
    }

    internal static string DescribeSource(IElement element) {
        string tag = element.TagName.ToLowerInvariant();
        if (!string.IsNullOrWhiteSpace(element.Id)) return tag + "#" + element.Id;
        string? className = element.GetAttribute("class");
        if (!string.IsNullOrWhiteSpace(className)) return tag + "." + className!.Trim().Replace(' ', '.');
        return tag;
    }

    private double ResolveFontSize(string value, double parentFontSize) {
        if (string.IsNullOrWhiteSpace(value)) return parentFontSize;
        string normalized = value.Trim().ToLowerInvariant();
        switch (normalized) {
            case "xx-small": return parentFontSize * 0.6D;
            case "x-small": return parentFontSize * 0.75D;
            case "small": return parentFontSize * 0.89D;
            case "medium": return _options.DefaultFontSize;
            case "large": return parentFontSize * 1.2D;
            case "x-large": return parentFontSize * 1.5D;
            case "xx-large": return parentFontSize * 2D;
            case "smaller": return parentFontSize * 0.8D;
            case "larger": return parentFontSize * 1.2D;
        }

        return HtmlRenderCssValues.TryLength(normalized, parentFontSize, parentFontSize, _options.DefaultFontSize, out double size) && size > 0D
            ? size
            : parentFontSize;
    }

    private static OfficeFontStyle ResolveFontStyle(string tag, HtmlComputedStyle computed) {
        OfficeFontStyle result = OfficeFontStyle.Regular;
        string weight = computed.GetValue("font-weight");
        bool heading = tag.Length == 2 && tag[0] == 'h' && tag[1] >= '1' && tag[1] <= '6';
        if (heading || tag == "b" || tag == "strong" || string.Equals(weight, "bold", StringComparison.OrdinalIgnoreCase) || TryFontWeight(weight, out int numericWeight) && numericWeight >= 600) {
            result |= OfficeFontStyle.Bold;
        }

        string style = computed.GetValue("font-style");
        if (tag == "i" || tag == "em" || style.StartsWith("italic", StringComparison.OrdinalIgnoreCase) || style.StartsWith("oblique", StringComparison.OrdinalIgnoreCase)) {
            result |= OfficeFontStyle.Italic;
        }

        string decoration = computed.GetValue("text-decoration-line");
        if (tag == "u" || decoration.IndexOf("underline", StringComparison.OrdinalIgnoreCase) >= 0) result |= OfficeFontStyle.Underline;
        if (tag == "s" || tag == "strike" || tag == "del" || decoration.IndexOf("line-through", StringComparison.OrdinalIgnoreCase) >= 0) result |= OfficeFontStyle.Strikethrough;
        return result;
    }

    private static bool TryFontWeight(string value, out int weight) => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out weight);

    private static double ResolveDefaultTagFontSize(string tag, double inherited) {
        switch (tag) {
            case "h1": return inherited * 2D;
            case "h2": return inherited * 1.5D;
            case "h3": return inherited * 1.17D;
            case "h4": return inherited;
            case "h5": return inherited * 0.83D;
            case "h6": return inherited * 0.67D;
            case "small": return inherited * 0.83D;
            case "big": return inherited * 1.17D;
            default: return inherited;
        }
    }

    private static string ResolveDisplay(string tag, string value) {
        if (!string.IsNullOrWhiteSpace(value)) return value.Trim().ToLowerInvariant();
        if (tag == "li") return "list-item";
        if (tag == "table") return "table";
        return IsDefaultBlockTag(tag) ? "block" : "inline";
    }

    private static string ResolvePseudoDisplay(string value) =>
        string.IsNullOrWhiteSpace(value) ? "inline" : value.Trim().ToLowerInvariant();

    private static string ResolveListStyleType(HtmlComputedStyle computed) {
        string type = computed.GetValue("list-style-type").Trim().ToLowerInvariant();
        if (type.Length > 0) return type;
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(computed.GetValue("list-style"))) {
            if (string.Equals(token, "none", StringComparison.OrdinalIgnoreCase)) return "none";
        }

        return string.Empty;
    }

    private static bool IsDefaultBlockTag(string tagName) {
        string tag = tagName.ToLowerInvariant();
        return tag == "html" || tag == "body" || tag == "address" || tag == "article" || tag == "aside" || tag == "blockquote"
            || tag == "details" || tag == "dialog" || tag == "div" || tag == "dl" || tag == "dt" || tag == "dd" || tag == "fieldset"
            || tag == "figcaption" || tag == "figure" || tag == "footer" || tag == "form" || tag == "h1" || tag == "h2" || tag == "h3"
            || tag == "h4" || tag == "h5" || tag == "h6" || tag == "header" || tag == "hr" || tag == "li" || tag == "main"
            || tag == "nav" || tag == "ol" || tag == "p" || tag == "pre" || tag == "section" || tag == "summary" || tag == "table"
            || tag == "ul";
    }

    internal static bool IsDefaultBlockElement(IElement element) => IsDefaultBlockTag(element.TagName);

    private static OfficeColor ResolveColor(string value, OfficeColor fallback) => HtmlRenderCssValues.TryColor(value, out OfficeColor color) ? color : fallback;

    private static OfficeTextAlignment ResolveAlignment(string value, string direction) {
        if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Center;
        if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Right;
        if (string.Equals(value, "left", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Left;
        bool rightToLeft = string.Equals(direction, "rtl", StringComparison.Ordinal);
        if (string.Equals(value, "end", StringComparison.OrdinalIgnoreCase)) return rightToLeft ? OfficeTextAlignment.Left : OfficeTextAlignment.Right;
        return rightToLeft ? OfficeTextAlignment.Right : OfficeTextAlignment.Left;
    }

    private double ResolveLineHeight(string value, double fontSize) {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "normal", StringComparison.OrdinalIgnoreCase)) return fontSize * _options.DefaultLineHeight;
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double multiplier) && multiplier > 0D) return fontSize * multiplier;
        return HtmlRenderCssValues.TryLength(value, fontSize, fontSize, _options.DefaultFontSize, out double lineHeight) && lineHeight > 0D
            ? lineHeight
            : fontSize * _options.DefaultLineHeight;
    }

    private static string ResolveSemanticRole(string tag) {
        if (tag.Length == 2 && tag[0] == 'h' && tag[1] >= '1' && tag[1] <= '6') return "heading-" + tag[1];
        if (tag == "p") return "paragraph";
        if (tag == "li") return "list-item";
        if (tag == "th") return "table-header";
        if (tag == "td") return "table-cell";
        if (tag == "figcaption") return "caption";
        return tag;
    }

    private static bool IsPreformatted(string tag, string whiteSpace) => tag == "pre" || whiteSpace == "pre" || whiteSpace == "pre-wrap" || whiteSpace == "break-spaces";

    private static void ApplyDefaultMargins(string tag, double fontSize, HtmlRenderBoxStyle style) {
        if (tag == "p" || tag == "pre" || tag == "blockquote" || tag == "table" || tag == "figure" || tag == "ul" || tag == "ol") {
            style.MarginBottom = fontSize;
        } else if (tag.Length == 2 && tag[0] == 'h' && tag[1] >= '1' && tag[1] <= '6') {
            style.MarginTop = fontSize * 0.67D;
            style.MarginBottom = fontSize * 0.67D;
        } else if (tag == "li") {
            style.MarginBottom = fontSize * 0.25D;
        }

        if (tag == "blockquote") {
            style.MarginLeft = fontSize * 2D;
            style.MarginRight = fontSize * 2D;
        }
    }

    private void ApplyBoxValues(HtmlComputedStyle computed, double reference, double fontSize, HtmlRenderBoxStyle style) {
        string margin = computed.GetValue("margin");
        ApplyAutoMargins(computed, margin, style);
        if (margin.Length > 0) HtmlRenderCssValues.ApplyBoxShorthand(margin, reference, fontSize, _options.DefaultFontSize, ref style.MarginTop, ref style.MarginRight, ref style.MarginBottom, ref style.MarginLeft);
        ApplyLength(computed.GetValue("margin-top"), reference, fontSize, ref style.MarginTop);
        ApplyLength(computed.GetValue("margin-right"), reference, fontSize, ref style.MarginRight);
        ApplyLength(computed.GetValue("margin-bottom"), reference, fontSize, ref style.MarginBottom);
        ApplyLength(computed.GetValue("margin-left"), reference, fontSize, ref style.MarginLeft);

        string padding = computed.GetValue("padding");
        if (padding.Length > 0) HtmlRenderCssValues.ApplyBoxShorthand(padding, reference, fontSize, _options.DefaultFontSize, ref style.PaddingTop, ref style.PaddingRight, ref style.PaddingBottom, ref style.PaddingLeft);
        ApplyLength(computed.GetValue("padding-top"), reference, fontSize, ref style.PaddingTop);
        ApplyLength(computed.GetValue("padding-right"), reference, fontSize, ref style.PaddingRight);
        ApplyLength(computed.GetValue("padding-bottom"), reference, fontSize, ref style.PaddingBottom);
        ApplyLength(computed.GetValue("padding-left"), reference, fontSize, ref style.PaddingLeft);

        ApplyBorderAndOutlinePaint(computed, reference, fontSize, style);
    }

    private static void ApplyAutoMargins(HtmlComputedStyle computed, string shorthand, HtmlRenderBoxStyle style) {
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(shorthand);
        string top = values.Count > 0 ? values[0] : string.Empty;
        string right = values.Count > 1 ? values[1] : top;
        string bottom = values.Count > 2 ? values[2] : top;
        string left = values.Count > 3 ? values[3] : right;
        style.MarginTopAuto = string.Equals(top, "auto", StringComparison.OrdinalIgnoreCase);
        style.MarginRightAuto = string.Equals(right, "auto", StringComparison.OrdinalIgnoreCase);
        style.MarginBottomAuto = string.Equals(bottom, "auto", StringComparison.OrdinalIgnoreCase);
        style.MarginLeftAuto = string.Equals(left, "auto", StringComparison.OrdinalIgnoreCase);
        OverrideAutoMargin(computed.GetValue("margin-top"), ref style.MarginTopAuto);
        OverrideAutoMargin(computed.GetValue("margin-right"), ref style.MarginRightAuto);
        OverrideAutoMargin(computed.GetValue("margin-bottom"), ref style.MarginBottomAuto);
        OverrideAutoMargin(computed.GetValue("margin-left"), ref style.MarginLeftAuto);
    }

    private static void OverrideAutoMargin(string value, ref bool target) {
        if (!string.IsNullOrWhiteSpace(value)) target = string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase);
    }

    private void ApplyDimensions(
        IElement element,
        HtmlComputedStyle computed,
        double reference,
        double fontSize,
        HtmlRenderBoxStyle? parent,
        HtmlRenderBoxStyle style,
        bool includeAttributes) {
        style.ExplicitWidth = ReadLength(computed.GetValue("width"), includeAttributes ? element.GetAttribute("width") : null, reference, fontSize);
        double? parentContentHeight = ResolveDefiniteContentHeight(parent);
        style.ExplicitHeight = ReadVerticalLength(computed.GetValue("height"), includeAttributes ? element.GetAttribute("height") : null, reference, parentContentHeight, fontSize);
        style.MinWidth = ReadLength(computed.GetValue("min-width"), null, reference, fontSize);
        style.MaxWidth = ReadLength(computed.GetValue("max-width"), null, reference, fontSize);
        style.MinHeight = ReadVerticalLength(computed.GetValue("min-height"), null, reference, parentContentHeight, fontSize);
        style.MaxHeight = ReadVerticalLength(computed.GetValue("max-height"), null, reference, parentContentHeight, fontSize);
    }

    private void ApplyPaint(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        string backgroundShorthand = computed.GetValue("background");
        string background = computed.GetValue("background-color");
        if (background.Length == 0) background = backgroundShorthand;
        if (HtmlRenderCssValues.TryColor(background, out OfficeColor backgroundColor)) style.BackgroundColor = backgroundColor;
        ApplyBackgroundLayers(computed, style, backgroundShorthand);
        ApplyOpacity(computed.GetValue("opacity"), style);
        style.Transform = NormalizeCssValue(computed.GetValue("transform"), "none");
        style.TransformOrigin = NormalizeCssValue(computed.GetValue("transform-origin"), "50% 50%");
        string boxShadow = NormalizeCssValue(computed.GetValue("box-shadow"), "none");
        if (!HtmlCssBoxShadowParser.TryParse(boxShadow, style.Font.Size, _options.DefaultFontSize, style.Color, out IReadOnlyList<HtmlCssBoxShadow> shadows)) {
            style.UnsupportedBoxShadow = boxShadow;
        } else {
            style.BoxShadowLayerCount = shadows.Count;
            style.BoxShadows = shadows.Take(_options.MaxBoxShadowLayers).ToArray();
        }
    }

    private static void ApplyOpacity(string value, HtmlRenderBoxStyle style) {
        if (string.IsNullOrWhiteSpace(value)) return;
        style.OpacityWasSpecified = true;
        string normalized = value.Trim().ToLowerInvariant();
        bool percentage = normalized.EndsWith("%", StringComparison.Ordinal);
        string numberText = percentage ? normalized.Substring(0, normalized.Length - 1) : normalized;
        if (!double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out double opacity)
            || double.IsNaN(opacity) || double.IsInfinity(opacity)) {
            style.UnsupportedOpacity = normalized;
            return;
        }
        if (percentage) opacity /= 100D;
        style.Opacity = Math.Max(0D, Math.Min(1D, opacity));
    }

    private void ApplyBackgroundLayers(HtmlComputedStyle computed, HtmlRenderBoxStyle style, string backgroundShorthand) {
        string backgroundImage = computed.GetValue("background-image");
        string sourceValue = backgroundImage.Length > 0 ? backgroundImage : backgroundShorthand;
        IReadOnlyList<string> sourceLayers = HtmlRenderCssValues.SplitTopLevelCommas(sourceValue);
        IReadOnlyList<string> positionLayers = HtmlRenderCssValues.SplitTopLevelCommas(computed.GetValue("background-position"));
        IReadOnlyList<string> repeatLayers = HtmlRenderCssValues.SplitTopLevelCommas(computed.GetValue("background-repeat"));
        IReadOnlyList<string> sizeLayers = HtmlRenderCssValues.SplitTopLevelCommas(computed.GetValue("background-size"));
        var layers = new List<HtmlRenderBackgroundLayer>();
        int declaredLayerCount = 0;
        bool hasDeclaredBackgroundImage = false;
        int unsupportedLayerCount = 0;
        int gradientStopLimitExceededCount = 0;
        for (int index = 0; index < sourceLayers.Count; index++) {
            string sourceLayer = sourceLayers[index];
            IReadOnlyList<string> urls = HtmlResourcePipeline.ExtractCssUrls(sourceLayer);
            bool isNone = string.Equals(sourceLayer.Trim(), "none", StringComparison.OrdinalIgnoreCase);
            bool hasGradientFunction = urls.Count == 0
                && sourceLayer.IndexOf("gradient(", StringComparison.OrdinalIgnoreCase) >= 0;
            if (urls.Count == 0 && !hasGradientFunction && !isNone) continue;

            declaredLayerCount++;
            if (!isNone) hasDeclaredBackgroundImage = true;
            if (declaredLayerCount > _options.MaxBackgroundImageLayers) continue;
            if (isNone) continue;
            string position = GetLayerValue(positionLayers, index, ExtractBackgroundPosition(sourceLayer), "0% 0%");
            string repeat = GetLayerValue(repeatLayers, index, ExtractBackgroundRepeat(sourceLayer), "repeat");
            string size = GetLayerValue(sizeLayers, index, ExtractBackgroundSize(sourceLayer), "auto");
            if (urls.Count == 0) {
                if (HtmlCssLinearGradientParser.TryParse(sourceLayer, _options.MaxGradientStops, out HtmlCssLinearGradientDefinition? linearGradient, out bool linearStopLimitExceeded)
                    && linearGradient != null) {
                    layers.Add(new HtmlRenderBackgroundLayer(linearGradient, position, repeat, size));
                    continue;
                }

                if (HtmlCssRadialGradientParser.TryParse(sourceLayer, _options.MaxGradientStops, out HtmlCssRadialGradientDefinition? radialGradient, out bool radialStopLimitExceeded)
                    && radialGradient != null) {
                    layers.Add(new HtmlRenderBackgroundLayer(radialGradient, position, repeat, size));
                    continue;
                }

                if (linearStopLimitExceeded || radialStopLimitExceeded) {
                    gradientStopLimitExceededCount++;
                } else {
                    unsupportedLayerCount++;
                }

                continue;
            }

            layers.Add(new HtmlRenderBackgroundLayer(urls[0], position, repeat, size));
        }

        style.BackgroundImageLayerCount = declaredLayerCount;
        style.HasDeclaredBackgroundImage = hasDeclaredBackgroundImage;
        style.UnsupportedBackgroundImageLayerCount = unsupportedLayerCount;
        style.GradientStopLimitExceededCount = gradientStopLimitExceededCount;
        style.BackgroundImageLayers = layers.AsReadOnly();
    }

    private static string GetLayerValue(IReadOnlyList<string> values, int index, string shorthandValue, string fallback) {
        if (values.Count > 0) {
            string value = values[index % values.Count].Trim();
            if (value.Length > 0) return value;
        }

        return shorthandValue.Length > 0 ? shorthandValue : fallback;
    }

    private static string ExtractBackgroundRepeat(string shorthand) {
        var values = new List<string>();
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(shorthand)) {
            string value = token.Trim().ToLowerInvariant();
            if (value == "repeat-x" || value == "repeat-y") {
                return value;
            }

            if (value == "repeat" || value == "no-repeat" || value == "space" || value == "round") {
                values.Add(value);
                if (values.Count == 2) break;
            }
        }

        return string.Join(" ", values);
    }

    private static string ExtractBackgroundSize(string shorthand) {
        int slash = FindTopLevelCharacter(shorthand, '/');
        if (slash < 0 || slash + 1 >= shorthand.Length) {
            return string.Empty;
        }

        var values = new List<string>();
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(shorthand.Substring(slash + 1))) {
            string value = token.Trim().TrimEnd(',');
            if (value == "cover" || value == "contain" || value == "auto" || LooksLikeBackgroundLength(value)) {
                values.Add(value);
                if (values.Count == 2) break;
            } else if (values.Count > 0) {
                break;
            }
        }

        return string.Join(" ", values);
    }

    private static string ExtractBackgroundPosition(string shorthand) {
        string beforeSize = shorthand;
        int slash = FindTopLevelCharacter(beforeSize, '/');
        if (slash >= 0) beforeSize = beforeSize.Substring(0, slash);
        var values = new List<string>();
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(beforeSize)) {
            string value = token.Trim().TrimEnd(',').ToLowerInvariant();
            if (value == "left" || value == "right" || value == "top" || value == "bottom" || value == "center" || LooksLikeBackgroundLength(value)) {
                values.Add(value);
                if (values.Count == 2) break;
            }
        }

        return string.Join(" ", values);
    }

    private static bool LooksLikeBackgroundLength(string value) =>
        value.EndsWith("%", StringComparison.Ordinal)
        || value.EndsWith("px", StringComparison.OrdinalIgnoreCase)
        || value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
        || value.EndsWith("in", StringComparison.OrdinalIgnoreCase)
        || value.EndsWith("cm", StringComparison.OrdinalIgnoreCase)
        || value.EndsWith("mm", StringComparison.OrdinalIgnoreCase)
        || value == "0";

    private static int FindTopLevelCharacter(string value, char target) {
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (quote != '\0') {
                if (current == quote && (index == 0 || value[index - 1] != '\\')) quote = '\0';
                continue;
            }

            if (current == '\'' || current == '"') {
                quote = current;
            } else if (current == '(') {
                depth++;
            } else if (current == ')' && depth > 0) {
                depth--;
            } else if (current == target && depth == 0) {
                return index;
            }
        }

        return -1;
    }

    private static void ApplyBreaks(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        string before = FirstNonEmpty(computed.GetValue("break-before"), computed.GetValue("page-break-before"));
        string after = FirstNonEmpty(computed.GetValue("break-after"), computed.GetValue("page-break-after"));
        string inside = FirstNonEmpty(computed.GetValue("break-inside"), computed.GetValue("page-break-inside"));
        style.BreakBefore = ResolvePageBreakTarget(before);
        style.BreakAfter = ResolvePageBreakTarget(after);
        style.AvoidBreakInside = string.Equals(inside, "avoid", StringComparison.OrdinalIgnoreCase) || string.Equals(inside, "avoid-page", StringComparison.OrdinalIgnoreCase);
        style.Orphans = ReadPositiveInteger(computed.GetValue("orphans"), style.Orphans);
        style.Widows = ReadPositiveInteger(computed.GetValue("widows"), style.Widows);
        style.PageName = ResolvePageName(computed.GetValue("page"));
    }

    private static void ApplyPositioning(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        style.Position = NormalizeCssValue(computed.GetValue("position"), "static");
        style.Top = NormalizeCssValue(computed.GetValue("top"), "auto");
        style.Right = NormalizeCssValue(computed.GetValue("right"), "auto");
        style.Bottom = NormalizeCssValue(computed.GetValue("bottom"), "auto");
        style.Left = NormalizeCssValue(computed.GetValue("left"), "auto");
        style.ZIndex = NormalizeCssValue(computed.GetValue("z-index"), "auto");
    }

    private void ApplyFlex(HtmlComputedStyle computed, double reference, double fontSize, HtmlRenderBoxStyle style) {
        style.FlexDirection = NormalizeCssValue(computed.GetValue("flex-direction"), "row");
        style.FlexWrap = NormalizeCssValue(computed.GetValue("flex-wrap"), "nowrap");
        ApplyFlexFlow(computed.GetValue("flex-flow"), style);
        style.JustifyContent = NormalizeCssValue(computed.GetValue("justify-content"), "normal");
        style.AlignItems = NormalizeCssValue(computed.GetValue("align-items"), "normal");
        style.AlignContent = NormalizeCssValue(computed.GetValue("align-content"), "normal");
        style.AlignSelf = NormalizeCssValue(computed.GetValue("align-self"), "auto");
        ApplyFlexShorthand(computed.GetValue("flex"), style);
        if (TryNonNegativeNumber(computed.GetValue("flex-grow"), out double grow)) style.FlexGrow = grow;
        if (TryNonNegativeNumber(computed.GetValue("flex-shrink"), out double shrink)) style.FlexShrink = shrink;
        string basis = computed.GetValue("flex-basis");
        if (!string.IsNullOrWhiteSpace(basis)) style.FlexBasis = basis.Trim().ToLowerInvariant();
        if (int.TryParse(computed.GetValue("order"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int order)) style.Order = order;
        ApplyGap(computed, reference, fontSize, style);
    }

    private void ApplyColumns(HtmlComputedStyle computed, double reference, double fontSize, HtmlRenderBoxStyle style) {
        string shorthand = computed.GetValue("columns");
        if (!string.IsNullOrWhiteSpace(shorthand)) {
            style.ColumnCount = null;
            style.ColumnWidth = null;
            foreach (string token in HtmlRenderCssValues.SplitWhitespace(shorthand)) {
                string normalized = token.Trim().ToLowerInvariant();
                if (normalized == "auto") continue;
                if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) && count > 0 && !style.ColumnCount.HasValue) {
                    style.ColumnCount = count;
                } else if (TryResolveColumnWidth(normalized, reference, fontSize, out double width) && !style.ColumnWidth.HasValue) {
                    style.ColumnWidth = width;
                } else {
                    style.UnsupportedColumns = shorthand.Trim();
                }
            }
        }

        string countValue = computed.GetValue("column-count");
        if (!string.IsNullOrWhiteSpace(countValue)) {
            string normalized = countValue.Trim().ToLowerInvariant();
            if (normalized == "auto") style.ColumnCount = null;
            else if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) && count > 0) style.ColumnCount = count;
            else style.UnsupportedColumns = "column-count=" + normalized;
        }
        string widthValue = computed.GetValue("column-width");
        if (!string.IsNullOrWhiteSpace(widthValue)) {
            string normalized = widthValue.Trim().ToLowerInvariant();
            if (normalized == "auto") style.ColumnWidth = null;
            else if (TryResolveColumnWidth(normalized, reference, fontSize, out double width)) style.ColumnWidth = width;
            else style.UnsupportedColumns = "column-width=" + normalized;
        }

        string fill = NormalizeCssValue(computed.GetValue("column-fill"), "balance");
        if (fill == "auto" || fill == "balance") style.ColumnFill = fill;
        else style.UnsupportedColumnFill = fill;
        string span = NormalizeCssValue(computed.GetValue("column-span"), "none");
        if (span == "none" || span == "all") style.ColumnSpan = span;
        else style.UnsupportedColumnSpan = span;
        ApplyColumnRule(computed, reference, fontSize, style);
    }

    private void ApplyColumnRule(HtmlComputedStyle computed, double reference, double fontSize, HtmlRenderBoxStyle style) {
        string shorthand = computed.GetValue("column-rule");
        if (!string.IsNullOrWhiteSpace(shorthand)) {
            foreach (string token in HtmlRenderCssValues.SplitWhitespace(shorthand)) {
                if (TryResolveColumnRuleWidth(token, reference, fontSize, out double width)) {
                    style.ColumnRuleWidth = width;
                } else if (TryResolveColumnRuleStyle(token, out string ruleStyle)) {
                    style.ColumnRuleStyle = ruleStyle;
                } else if (TryResolveColumnRuleColor(token, style.Color, out OfficeColor color)) {
                    style.ColumnRuleColor = color;
                } else {
                    style.UnsupportedColumnRule = shorthand.Trim();
                }
            }
        }

        string widthValue = computed.GetValue("column-rule-width");
        if (!string.IsNullOrWhiteSpace(widthValue)) {
            if (TryResolveColumnRuleWidth(widthValue, reference, fontSize, out double width)) style.ColumnRuleWidth = width;
            else style.UnsupportedColumnRule = "column-rule-width=" + widthValue.Trim();
        }
        string styleValue = computed.GetValue("column-rule-style");
        if (!string.IsNullOrWhiteSpace(styleValue)) {
            if (TryResolveColumnRuleStyle(styleValue, out string ruleStyle)) style.ColumnRuleStyle = ruleStyle;
            else style.UnsupportedColumnRule = "column-rule-style=" + styleValue.Trim();
        }
        string colorValue = computed.GetValue("column-rule-color");
        if (!string.IsNullOrWhiteSpace(colorValue)) {
            if (TryResolveColumnRuleColor(colorValue, style.Color, out OfficeColor color)) style.ColumnRuleColor = color;
            else style.UnsupportedColumnRule = "column-rule-color=" + colorValue.Trim();
        }
    }

    private bool TryResolveColumnRuleWidth(string value, double reference, double fontSize, out double width) {
        width = 0D;
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized == "thin") {
            width = 1D;
            return true;
        }
        if (normalized == "medium") {
            width = 3D;
            return true;
        }
        if (normalized == "thick") {
            width = 5D;
            return true;
        }
        return (normalized == "0" || HasSupportedColumnLengthUnit(normalized))
            && HtmlRenderCssValues.TryLength(normalized, reference, fontSize, _options.DefaultFontSize, out width)
            && width >= 0D;
    }

    private static bool TryResolveColumnRuleStyle(string value, out string style) {
        style = value.Trim().ToLowerInvariant();
        return style == "none" || style == "hidden" || style == "solid" || style == "dashed" || style == "dotted" || style == "double";
    }

    private static bool TryResolveColumnRuleColor(string value, OfficeColor currentColor, out OfficeColor color) {
        if (string.Equals(value.Trim(), "currentcolor", StringComparison.OrdinalIgnoreCase)) {
            color = currentColor;
            return true;
        }
        return HtmlRenderCssValues.TryColor(value, out color);
    }

    private bool TryResolveColumnWidth(string value, double reference, double fontSize, out double width) {
        width = 0D;
        return HasSupportedColumnLengthUnit(value)
            && HtmlRenderCssValues.TryLength(value, reference, fontSize, _options.DefaultFontSize, out width)
            && width > 0D;
    }

    private static bool HasSupportedColumnLengthUnit(string value) {
        string normalized = value.Trim().ToLowerInvariant();
        int unitStart = 0;
        while (unitStart < normalized.Length && (char.IsDigit(normalized[unitStart]) || normalized[unitStart] == '.' || normalized[unitStart] == '+' || normalized[unitStart] == '-')) unitStart++;
        if (unitStart == 0 || unitStart == normalized.Length) return false;
        if (!double.TryParse(normalized.Substring(0, unitStart), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            || double.IsNaN(number) || double.IsInfinity(number)) return false;
        string unit = normalized.Substring(unitStart);
        return unit == "px" || unit == "pt" || unit == "pc" || unit == "in" || unit == "cm" || unit == "mm" || unit == "q" || unit == "em" || unit == "rem";
    }

    private static void ApplyGrid(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        style.GridTemplateColumns = NormalizeCssValue(computed.GetValue("grid-template-columns"), "none");
        style.GridTemplateRows = NormalizeCssValue(computed.GetValue("grid-template-rows"), "none");
        style.GridTemplateAreas = NormalizeCssValue(computed.GetValue("grid-template-areas"), "none");
        style.GridAutoColumns = NormalizeCssValue(computed.GetValue("grid-auto-columns"), "auto");
        style.GridAutoRows = NormalizeCssValue(computed.GetValue("grid-auto-rows"), "auto");
        style.GridAutoFlow = NormalizeCssValue(computed.GetValue("grid-auto-flow"), "row");
        style.JustifyItems = NormalizeCssValue(computed.GetValue("justify-items"), "normal");
        style.JustifySelf = NormalizeCssValue(computed.GetValue("justify-self"), "auto");
        ApplyGridPair(computed.GetValue("grid-column"), ref style.GridColumnStart, ref style.GridColumnEnd);
        ApplyGridPair(computed.GetValue("grid-row"), ref style.GridRowStart, ref style.GridRowEnd);
        style.GridArea = NormalizeCssValue(computed.GetValue("grid-area"), "auto");
        ApplyGridArea(computed.GetValue("grid-area"), style);
        OverrideGridValue(computed.GetValue("grid-column-start"), ref style.GridColumnStart);
        OverrideGridValue(computed.GetValue("grid-column-end"), ref style.GridColumnEnd);
        OverrideGridValue(computed.GetValue("grid-row-start"), ref style.GridRowStart);
        OverrideGridValue(computed.GetValue("grid-row-end"), ref style.GridRowEnd);
        ApplyPlacePair(computed.GetValue("place-items"), ref style.AlignItems, ref style.JustifyItems);
        ApplyPlacePair(computed.GetValue("place-self"), ref style.AlignSelf, ref style.JustifySelf);
        ApplyPlacePair(computed.GetValue("place-content"), ref style.AlignContent, ref style.JustifyContent);
    }

    private static void ApplyGridPair(string value, ref string start, ref string end) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitTopLevel(value, '/');
        if (parts.Count > 0 && !string.IsNullOrWhiteSpace(parts[0])) start = parts[0].Trim().ToLowerInvariant();
        if (parts.Count > 1 && !string.IsNullOrWhiteSpace(parts[1])) end = parts[1].Trim().ToLowerInvariant();
    }

    private static void ApplyGridArea(string value, HtmlRenderBoxStyle style) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitTopLevel(value, '/');
        if (parts.Count > 0 && !string.IsNullOrWhiteSpace(parts[0])) style.GridRowStart = parts[0].Trim().ToLowerInvariant();
        if (parts.Count > 1 && !string.IsNullOrWhiteSpace(parts[1])) style.GridColumnStart = parts[1].Trim().ToLowerInvariant();
        if (parts.Count > 2 && !string.IsNullOrWhiteSpace(parts[2])) style.GridRowEnd = parts[2].Trim().ToLowerInvariant();
        if (parts.Count > 3 && !string.IsNullOrWhiteSpace(parts[3])) style.GridColumnEnd = parts[3].Trim().ToLowerInvariant();
    }

    private static void ApplyPlacePair(string value, ref string first, ref string second) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        if (parts.Count == 0) return;
        first = parts[0].Trim().ToLowerInvariant();
        second = (parts.Count > 1 ? parts[1] : parts[0]).Trim().ToLowerInvariant();
    }

    private static void OverrideGridValue(string value, ref string target) {
        if (!string.IsNullOrWhiteSpace(value)) target = value.Trim().ToLowerInvariant();
    }

    private static void ApplyFlexFlow(string value, HtmlRenderBoxStyle style) {
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(value)) {
            string normalized = token.Trim().ToLowerInvariant();
            if (normalized == "row" || normalized == "row-reverse" || normalized == "column" || normalized == "column-reverse") style.FlexDirection = normalized;
            else if (normalized == "nowrap" || normalized == "wrap" || normalized == "wrap-reverse") style.FlexWrap = normalized;
        }
    }

    private static void ApplyFlexShorthand(string value, HtmlRenderBoxStyle style) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0) return;
        if (normalized == "none") {
            style.FlexGrow = 0D;
            style.FlexShrink = 0D;
            style.FlexBasis = "auto";
            return;
        }

        if (normalized == "auto") {
            style.FlexGrow = 1D;
            style.FlexShrink = 1D;
            style.FlexBasis = "auto";
            return;
        }

        if (normalized == "initial") return;
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(normalized);
        if (parts.Count == 0 || !TryNonNegativeNumber(parts[0], out double grow)) return;
        style.FlexGrow = grow;
        style.FlexBasis = "0%";
        if (parts.Count == 1) return;
        if (TryNonNegativeNumber(parts[1], out double shrink)) {
            style.FlexShrink = shrink;
            if (parts.Count > 2) style.FlexBasis = parts[2];
        } else {
            style.FlexBasis = parts[1];
        }
    }

    private void ApplyGap(HtmlComputedStyle computed, double reference, double fontSize, HtmlRenderBoxStyle style) {
        IReadOnlyList<string> gap = HtmlRenderCssValues.SplitWhitespace(computed.GetValue("gap"));
        string row = gap.Count > 0 ? gap[0] : string.Empty;
        string column = gap.Count > 1 ? gap[1] : row;
        if (!string.IsNullOrWhiteSpace(computed.GetValue("row-gap"))) row = computed.GetValue("row-gap");
        if (!string.IsNullOrWhiteSpace(computed.GetValue("column-gap"))) column = computed.GetValue("column-gap");
        style.ColumnGapWasSpecified = !string.IsNullOrWhiteSpace(column) && !string.Equals(column.Trim(), "normal", StringComparison.OrdinalIgnoreCase);
        style.RowGap = ResolveGap(row, reference, fontSize, out bool rowUnsupported);
        style.ColumnGap = ResolveGap(column, reference, fontSize, out bool columnUnsupported);
        if (rowUnsupported) style.UnsupportedRowGap = row.Trim();
        if (columnUnsupported) style.UnsupportedColumnGap = column.Trim();
    }

    private double ResolveGap(string value, double reference, double fontSize, out bool unsupported) {
        unsupported = false;
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "normal", StringComparison.OrdinalIgnoreCase)) return 0D;
        if (HtmlRenderCssValues.TryLength(value, reference, fontSize, _options.DefaultFontSize, out double resolved) && resolved >= 0D) return resolved;
        unsupported = true;
        return 0D;
    }

    private static bool TryNonNegativeNumber(string value, out double result) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
        && !double.IsNaN(result)
        && !double.IsInfinity(result)
        && result >= 0D;

    private static string NormalizeCssValue(string value, string fallback) =>
        string.IsNullOrWhiteSpace(value) ? fallback : value.Trim().ToLowerInvariant();

    private static string? ResolvePageName(string value) {
        string normalized = value.Trim();
        return normalized.Length == 0 || string.Equals(normalized, "auto", StringComparison.OrdinalIgnoreCase)
            ? null
            : normalized;
    }

    private static int ReadPositiveInteger(string value, int fallback) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed > 0
            ? parsed
            : fallback;

    private void ApplyLength(string value, double reference, double fontSize, ref double target) {
        if (HtmlRenderCssValues.TryLength(value, reference, fontSize, _options.DefaultFontSize, out double parsed)) target = Math.Max(0D, parsed);
    }

    private double? ReadLength(string cssValue, string? attributeValue, double reference, double fontSize) {
        string value = cssValue.Length > 0 ? cssValue : attributeValue ?? string.Empty;
        return HtmlRenderCssValues.TryLength(value, reference, fontSize, _options.DefaultFontSize, out double parsed) && parsed >= 0D ? parsed : null;
    }

    private double? ReadVerticalLength(
        string cssValue,
        string? attributeValue,
        double fallbackReference,
        double? parentContentHeight,
        double fontSize) {
        string value = cssValue.Length > 0 ? cssValue : attributeValue ?? string.Empty;
        string normalized = value.Trim();
        if (normalized.EndsWith("%", StringComparison.Ordinal)) {
            if (!parentContentHeight.HasValue
                || !double.TryParse(
                    normalized.Substring(0, normalized.Length - 1),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out double percentage)
                || percentage < 0D
                || double.IsNaN(percentage)
                || double.IsInfinity(percentage)) {
                return null;
            }
            return parentContentHeight.Value * percentage / 100D;
        }

        return ReadLength(cssValue, attributeValue, fallbackReference, fontSize);
    }

    private static double? ResolveDefiniteContentHeight(HtmlRenderBoxStyle? style) {
        if (style == null || !style.ExplicitHeight.HasValue) return null;
        return style.BorderBox
            ? Math.Max(0D, style.ExplicitHeight.Value - style.VerticalInsets)
            : style.ExplicitHeight.Value;
    }

    private static HtmlPageBreakTarget ResolvePageBreakTarget(string value) {
        if (value == "left" || value == "verso") return HtmlPageBreakTarget.Left;
        if (value == "right" || value == "recto") return HtmlPageBreakTarget.Right;
        if (value == "page" || value == "always") return HtmlPageBreakTarget.Page;
        return HtmlPageBreakTarget.None;
    }

    private static string FirstNonEmpty(string first, string second) => first.Length > 0 ? first.Trim().ToLowerInvariant() : second.Trim().ToLowerInvariant();

    private static string FirstNonEmpty(string first, string second, string third) =>
        FirstNonEmpty(first, FirstNonEmpty(second, third));
}
