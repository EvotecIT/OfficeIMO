using System.Globalization;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlRenderStyleResolver {
    private readonly IReadOnlyDictionary<IElement, HtmlComputedStyle> _computedStyles;
    private readonly HtmlRenderOptions _options;

    internal HtmlRenderStyleResolver(IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles, HtmlRenderOptions options) {
        _computedStyles = computedStyles;
        _options = options;
    }

    internal HtmlRenderBoxStyle Resolve(IElement element, double containingWidth, HtmlRenderBoxStyle? parent = null) {
        HtmlComputedStyle computed = _computedStyles.TryGetValue(element, out HtmlComputedStyle? found)
            ? found
            : new HtmlComputedStyle(new Dictionary<string, string>());
        string tag = element.TagName.ToLowerInvariant();
        double parentFontSize = parent?.Font.Size ?? _options.DefaultFontSize;
        string fontSizeValue = computed.GetValue("font-size");
        double fontSize = string.IsNullOrWhiteSpace(fontSizeValue)
            ? ResolveDefaultTagFontSize(tag, parentFontSize)
            : ResolveFontSize(fontSizeValue, parentFontSize);
        OfficeFontStyle fontStyle = ResolveFontStyle(tag, computed);
        string defaultFamily = tag == "code" || tag == "pre" || tag == "kbd" || tag == "samp"
            ? "Consolas"
            : parent?.Font.FamilyName ?? _options.DefaultFontFamily;
        string family = HtmlRenderCssValues.FontFamilyList(computed.GetValue("font-family"), defaultFamily);

        var style = new HtmlRenderBoxStyle {
            Display = ResolveDisplay(tag, computed.GetValue("display")),
            Font = new OfficeFontInfo(family, fontSize, fontStyle),
            Color = ResolveColor(computed.GetValue("color"), parent?.Color ?? OfficeColor.Black),
            Alignment = ResolveAlignment(computed.GetValue("text-align"), parent?.Alignment ?? OfficeTextAlignment.Left),
            LineHeight = ResolveLineHeight(computed.GetValue("line-height"), fontSize),
            SemanticRole = ResolveSemanticRole(tag),
            PreserveWhitespace = IsPreformatted(tag, computed.GetValue("white-space")),
            TextTransform = string.IsNullOrWhiteSpace(computed.GetValue("text-transform")) ? parent?.TextTransform ?? "none" : computed.GetValue("text-transform").Trim().ToLowerInvariant(),
            BorderBox = string.Equals(computed.GetValue("box-sizing"), "border-box", StringComparison.OrdinalIgnoreCase)
        };

        ApplyDefaultMargins(tag, fontSize, style);
        ApplyBoxValues(computed, containingWidth, fontSize, style);
        ApplyDimensions(element, computed, containingWidth, fontSize, style);
        ApplyPaint(computed, style);
        ApplyBreaks(computed, style);
        return style;
    }

    internal static bool IsBlockElement(IElement element, HtmlRenderBoxStyle style) {
        string display = style.Display;
        if (display == "none") return false;
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

    private static bool IsDefaultBlockTag(string tagName) {
        string tag = tagName.ToLowerInvariant();
        return tag == "html" || tag == "body" || tag == "address" || tag == "article" || tag == "aside" || tag == "blockquote"
            || tag == "details" || tag == "dialog" || tag == "div" || tag == "dl" || tag == "dt" || tag == "dd" || tag == "fieldset"
            || tag == "figcaption" || tag == "figure" || tag == "footer" || tag == "form" || tag == "h1" || tag == "h2" || tag == "h3"
            || tag == "h4" || tag == "h5" || tag == "h6" || tag == "header" || tag == "hr" || tag == "li" || tag == "main"
            || tag == "nav" || tag == "ol" || tag == "p" || tag == "pre" || tag == "section" || tag == "summary" || tag == "table"
            || tag == "ul" || tag == "img";
    }

    private static OfficeColor ResolveColor(string value, OfficeColor fallback) => HtmlRenderCssValues.TryColor(value, out OfficeColor color) ? color : fallback;

    private static OfficeTextAlignment ResolveAlignment(string value, OfficeTextAlignment fallback) {
        if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Center;
        if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "end", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Right;
        if (string.Equals(value, "left", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "start", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Left;
        return fallback;
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

        string border = computed.GetValue("border");
        string borderWidth = computed.GetValue("border-width");
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(borderWidth.Length > 0 ? borderWidth : border)) {
            if (HtmlRenderCssValues.TryLength(token, reference, fontSize, _options.DefaultFontSize, out double width) && width >= 0D) {
                style.BorderWidth = width;
                break;
            }
        }

        string borderColor = computed.GetValue("border-color");
        if (HtmlRenderCssValues.TryColor(borderColor.Length > 0 ? borderColor : border, out OfficeColor parsedBorderColor)) style.BorderColor = parsedBorderColor;
    }

    private void ApplyDimensions(IElement element, HtmlComputedStyle computed, double reference, double fontSize, HtmlRenderBoxStyle style) {
        style.ExplicitWidth = ReadLength(computed.GetValue("width"), element.GetAttribute("width"), reference, fontSize);
        style.ExplicitHeight = ReadLength(computed.GetValue("height"), element.GetAttribute("height"), reference, fontSize);
        style.MinWidth = ReadLength(computed.GetValue("min-width"), null, reference, fontSize);
        style.MaxWidth = ReadLength(computed.GetValue("max-width"), null, reference, fontSize);
        style.MinHeight = ReadLength(computed.GetValue("min-height"), null, reference, fontSize);
        style.MaxHeight = ReadLength(computed.GetValue("max-height"), null, reference, fontSize);
    }

    private void ApplyPaint(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        string backgroundShorthand = computed.GetValue("background");
        string background = computed.GetValue("background-color");
        if (background.Length == 0) background = backgroundShorthand;
        if (HtmlRenderCssValues.TryColor(background, out OfficeColor backgroundColor)) style.BackgroundColor = backgroundColor;
        ApplyBackgroundLayers(computed, style, backgroundShorthand);
        if (double.TryParse(computed.GetValue("opacity"), NumberStyles.Float, CultureInfo.InvariantCulture, out double opacity)) {
            style.Opacity = Math.Max(0D, Math.Min(1D, opacity));
            style.Color = HtmlRenderCssValues.ApplyOpacity(style.Color, style.Opacity);
            style.BorderColor = HtmlRenderCssValues.ApplyOpacity(style.BorderColor, style.Opacity);
            if (style.BackgroundColor.HasValue) style.BackgroundColor = HtmlRenderCssValues.ApplyOpacity(style.BackgroundColor.Value, style.Opacity);
        }
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
        int unsupportedLayerCount = 0;
        int gradientStopLimitExceededCount = 0;
        for (int index = 0; index < sourceLayers.Count; index++) {
            string sourceLayer = sourceLayers[index];
            IReadOnlyList<string> urls = HtmlResourcePipeline.ExtractCssUrls(sourceLayer);
            bool hasGradientFunction = urls.Count == 0
                && sourceLayer.IndexOf("gradient(", StringComparison.OrdinalIgnoreCase) >= 0;
            if (urls.Count == 0 && !hasGradientFunction) continue;

            declaredLayerCount++;
            if (declaredLayerCount > _options.MaxBackgroundImageLayers) continue;
            string position = GetLayerValue(positionLayers, index, ExtractBackgroundPosition(sourceLayer), "0% 0%");
            string repeat = GetLayerValue(repeatLayers, index, ExtractBackgroundRepeat(sourceLayer), "repeat");
            string size = GetLayerValue(sizeLayers, index, ExtractBackgroundSize(sourceLayer), "auto");
            if (urls.Count == 0) {
                if (HtmlCssLinearGradientParser.TryParse(sourceLayer, _options.MaxGradientStops, out OfficeLinearGradient? gradient, out bool stopLimitExceeded)
                    && gradient != null) {
                    layers.Add(new HtmlRenderBackgroundLayer(gradient, position, repeat, size));
                } else if (stopLimitExceeded) {
                    gradientStopLimitExceededCount++;
                } else {
                    unsupportedLayerCount++;
                }

                continue;
            }

            layers.Add(new HtmlRenderBackgroundLayer(urls[0], position, repeat, size));
        }

        style.BackgroundImageLayerCount = declaredLayerCount;
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
