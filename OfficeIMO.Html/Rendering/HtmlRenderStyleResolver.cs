using System.Globalization;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HtmlComputedStyleSet _computedStyles;
    private readonly HtmlRenderOptions _options;
    private readonly HtmlDiagnosticReport _diagnostics;
    private readonly Dictionary<IElement, HashSet<string>> _reportedUnsupportedColors = new Dictionary<IElement, HashSet<string>>();
    private readonly HashSet<IElement> _reportedSmallCapsApproximations = new HashSet<IElement>();
    private double _viewportWidth;
    private double _viewportHeight;
    private double _activeContainerWidth = double.NaN;
    private double _activeContainerHeight = double.NaN;

    internal HtmlRenderStyleResolver(HtmlComputedStyleSet computedStyles, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        _computedStyles = computedStyles;
        _options = options;
        _diagnostics = diagnostics;
        _viewportWidth = options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth;
        _viewportHeight = options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D;
    }

    internal void SetViewport(double width, double height) {
        _viewportWidth = width;
        _viewportHeight = height;
    }

    private bool TryResolveLength(string? value, double reference, double fontSize, double rootFontSize, out double result) =>
        HtmlRenderCssValues.TryLength(
            value,
            reference,
            fontSize,
            rootFontSize,
            _viewportWidth,
            _viewportHeight,
            _activeContainerWidth,
            _activeContainerHeight,
            out result);

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

        string semanticRole = kind switch {
            HtmlPseudoElementKind.Before => "generated-before",
            HtmlPseudoElementKind.After => "generated-after",
            _ => "list-marker"
        };
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
        double viewportWidth = _viewportWidth;
        double viewportHeight = _viewportHeight;
        _activeContainerWidth = parent?.ContainerType == "inline-size" || parent?.ContainerType == "size"
            ? containingWidth
            : parent?.ContainerUnitWidth ?? viewportWidth;
        _activeContainerHeight = parent?.ContainerType == "size"
            ? ResolveDefiniteContentHeight(parent) ?? viewportHeight
            : parent?.ContainerUnitHeight ?? viewportHeight;
        double parentFontSize = parent?.Font.Size ?? _options.DefaultFontSize;
        string fontSizeValue = computed.GetValue("font-size");
        double fontSize = computed.IsInheritedValue("font-size")
            ? parentFontSize
            : string.IsNullOrWhiteSpace(fontSizeValue)
            ? (pseudoElement ? parentFontSize : ResolveDefaultTagFontSize(tag, parentFontSize))
            : ResolveFontSize(fontSizeValue, parentFontSize);
        OfficeFontStyle fontStyle = ResolveFontStyle(pseudoElement ? string.Empty : tag, computed);
        OfficeTextDecorationStyle decorationStyle = ResolveTextDecorationStyle(computed.GetValue("text-decoration-style"));
        string defaultFamily = !pseudoElement && (tag == "code" || tag == "pre" || tag == "kbd" || tag == "samp")
            ? "Consolas"
            : parent?.Font.FamilyName ?? _options.DefaultFontFamily;
        string family = HtmlRenderCssValues.FontFamilyList(computed.GetValue("font-family"), defaultFamily);
        string direction = ResolveDirection(computed.GetValue("direction"), parent?.Direction);
        string language = ResolveLanguage(element, parent?.Language);

        string fontVariant = string.IsNullOrWhiteSpace(computed.GetValue("font-variant"))
            ? parent?.FontVariant ?? "normal"
            : computed.GetValue("font-variant").Trim().ToLowerInvariant();
        string fontVariantCaps = string.IsNullOrWhiteSpace(computed.GetValue("font-variant-caps"))
            ? fontVariant
            : computed.GetValue("font-variant-caps").Trim().ToLowerInvariant();
        string textTransform = string.IsNullOrWhiteSpace(computed.GetValue("text-transform"))
            ? parent?.TextTransform ?? "none"
            : computed.GetValue("text-transform").Trim().ToLowerInvariant();
        bool approximateSmallCaps = fontVariantCaps.IndexOf("small-caps", StringComparison.OrdinalIgnoreCase) >= 0;

        int baselineLevel = ResolveTextBaselineLevel(
            pseudoElement ? string.Empty : tag,
            computed.GetValue("vertical-align"),
            parent?.BaselineLevel ?? 0);
        (double baselineScale, double baselineOffset) = ResolveTextBaselineGeometry(
            pseudoElement ? string.Empty : tag,
            computed.GetValue("vertical-align"),
            parent?.BaselineScale ?? 1D,
            parent?.BaselineOffset ?? 0D,
            parent?.Font.Size ?? fontSize,
            parent?.LineHeight ?? fontSize * 1.2D);
        if (baselineLevel == 0 && Math.Abs(baselineOffset) > 0.000001D) baselineLevel = baselineOffset < 0D ? 1 : -1;
        OfficeColor color = ResolveColor(element, computed.GetValue("color"), parent?.Color ?? OfficeColor.Black, pseudoElement, "color");
        var style = new HtmlRenderBoxStyle {
            Display = pseudoElement ? ResolvePseudoDisplay(computed.GetValue("display")) : ResolveDisplay(element, computed.GetValue("display")),
            DisplayWasSpecified = !string.IsNullOrWhiteSpace(computed.GetValue("display")),
            PaintVisible = ResolvePaintVisibility(computed.GetValue("visibility"), parent),
            Font = new OfficeFontInfo(family, fontSize, fontStyle),
            UnderlineStyle = (fontStyle & OfficeFontStyle.Underline) == OfficeFontStyle.Underline
                ? decorationStyle
                : OfficeTextDecorationStyle.None,
            StrikethroughStyle = (fontStyle & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough
                ? decorationStyle
                : OfficeTextDecorationStyle.None,
            Baseline = baselineOffset switch {
                < 0D => OfficeTextBaseline.Superscript,
                > 0D => OfficeTextBaseline.Subscript,
                _ => OfficeTextBaseline.Normal
            },
            BaselineLevel = baselineLevel,
            BaselineScale = baselineScale,
            BaselineOffset = baselineOffset,
            Color = color,
            DecorationColor = ResolveColor(element, computed.GetValue("text-decoration-color"), color, pseudoElement, "text-decoration-color"),
            Alignment = ResolveAlignment(computed.GetValue("text-align"), direction, parent?.Alignment),
            LineHeight = ResolveLineHeight(computed.GetValue("line-height"), fontSize),
            LetterSpacing = ResolveTextSpacing(computed.GetValue("letter-spacing"), fontSize, parent?.LetterSpacing ?? 0D),
            WordSpacing = ResolveTextSpacing(computed.GetValue("word-spacing"), fontSize, parent?.WordSpacing ?? 0D),
            SemanticRole = pseudoElement ? pseudoSemanticRole : ResolveSemanticRole(tag),
            PreserveWhitespace = IsPreformatted(pseudoElement ? string.Empty : tag, computed.GetValue("white-space")),
            BreakSpaces = string.Equals(computed.GetValue("white-space"), "break-spaces", StringComparison.OrdinalIgnoreCase),
            PreventTextWrapping = PreventsTextWrapping(pseudoElement ? string.Empty : tag, computed.GetValue("white-space")),
            TextOverflow = ResolveTextOverflow(computed.GetValue("text-overflow")),
            LineClamp = ResolveLineClamp(computed),
            ListStyleType = ResolveListStyleType(computed),
            ListStylePosition = ResolveListStylePosition(computed),
            ListStyleImage = ResolveListStyleImage(computed),
            FontVariant = fontVariantCaps,
            TextTransform = textTransform,
            ApproximateSmallCaps = approximateSmallCaps,
            Language = language,
            Direction = direction,
            OverflowWrap = ResolveOverflowWrap(computed.GetValue("overflow-wrap"), parent?.OverflowWrap),
            WordBreak = ResolveWordBreak(computed.GetValue("word-break"), parent?.WordBreak),
            Hyphens = ResolveHyphens(computed.GetValue("hyphens"), parent?.Hyphens),
            HyphenateCharacter = ResolveHyphenateCharacter(computed.GetValue("hyphenate-character"), parent?.HyphenateCharacter),
            HyphenateLimitLines = ResolveHyphenateLimitLines(computed.GetValue("hyphenate-limit-lines"), parent?.HyphenateLimitLines),
            HyphenateLimitLast = ResolveHyphenateLimitLast(computed.GetValue("hyphenate-limit-last"), parent?.HyphenateLimitLast),
            BorderBox = string.Equals(computed.GetValue("box-sizing"), "border-box", StringComparison.OrdinalIgnoreCase)
        };
        if (approximateSmallCaps && _reportedSmallCapsApproximations.Add(element)) {
            _diagnostics.Add(
                "OfficeIMO.Html.Renderer",
                HtmlRenderDiagnosticCodes.FontVariantApproximated,
                "CSS small-caps used an uppercase managed-rendering approximation because synthetic small-cap glyph sizing is not available.",
                HtmlDiagnosticSeverity.Warning,
                DescribeSource(element),
                "font-variant-caps=" + fontVariantCaps,
                OfficeConversionLossKind.Approximation);
        }
        ResolveTabSize(
            computed.GetValue("tab-size"),
            computed.IsInheritedValue("tab-size"),
            computed.IsResetValue("tab-size"),
            parent?.TabSize ?? 8D,
            parent?.TabSizeIsLength ?? false,
            fontSize,
            out style.TabSize,
            out style.TabSizeIsLength);
        ResolveHyphenateLimitChars(computed.GetValue("hyphenate-limit-chars"), parent, style);
        style.HyphenateLimitZone = ResolveHyphenateLimitZone(computed.GetValue("hyphenate-limit-zone"), containingWidth, fontSize, parent?.HyphenateLimitZone ?? 0D);
        style.ContainerType = ResolveContainerTypeValue(computed);
        style.ContainerUnitWidth = _activeContainerWidth;
        style.ContainerUnitHeight = _activeContainerHeight;

        if (!pseudoElement) ApplyDefaultMargins(tag, fontSize, style);
        ApplyBoxValues(computed, containingWidth, fontSize, style);
        ApplyDimensions(element, computed, containingWidth, fontSize, parent, style, !pseudoElement);
        ApplyReplacedElementValues(computed, fontSize, style);
        ApplyPaint(element, computed, style, pseudoElement);
        if (style.OutlineColorInvert) {
            OfficeColor backdrop = style.BackgroundColor ?? parent?.BackgroundColor ?? OfficeColor.White;
            style.OutlineColor = OfficeColor.FromRgba((byte)(255 - backdrop.R), (byte)(255 - backdrop.G), (byte)(255 - backdrop.B), backdrop.A);
        }
        ApplyOverflow(computed, style);
        ApplyFloat(computed, style);
        ApplyPositioning(computed, style);
        ApplyFlex(computed, containingWidth, fontSize, style);
        ApplyColumns(computed, containingWidth, fontSize, style);
        ApplyGrid(computed, style);
        ApplyTable(computed, style);
        ApplyBreaks(computed, style);
        ApplyPdfSemanticTag(computed.GetValue("-officeimo-pdf-tag-type"), style);
        ApplyBookmark(computed, style);
        style.StringSet = computed.GetValue("string-set").Trim();
        return style;
    }

    private static string ResolveContainerTypeValue(HtmlComputedStyle computed) {
        string value = computed.GetValue("container-type").Trim().ToLowerInvariant();
        if (value == "size" || value == "inline-size") return value;
        string shorthand = computed.GetValue("container");
        int slash = shorthand.IndexOf('/');
        if (slash >= 0) {
            value = shorthand.Substring(slash + 1).Trim().ToLowerInvariant();
            if (value == "size" || value == "inline-size") return value;
        }
        return "normal";
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

    private static string ResolveHyphens(string value, string? inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited ?? "manual";
        return normalized == "none" || normalized == "manual" || normalized == "auto" ? normalized : "manual";
    }

    private static string ResolveHyphenateCharacter(string value, string? inherited) {
        string normalized = value.Trim();
        if (normalized.Length == 0 || string.Equals(normalized, "inherit", StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "unset", StringComparison.OrdinalIgnoreCase)) return inherited ?? "-";
        if (string.Equals(normalized, "auto", StringComparison.OrdinalIgnoreCase)) return "-";
        if (normalized.Length >= 2 && (normalized[0] == '\'' && normalized[normalized.Length - 1] == '\'' || normalized[0] == '"' && normalized[normalized.Length - 1] == '"')) {
            string decoded = HtmlCssEscapeDecoder.Decode(normalized.Substring(1, normalized.Length - 2));
            return decoded.Length <= 8 ? decoded : inherited ?? "-";
        }
        return inherited ?? "-";
    }

    private static void ResolveHyphenateLimitChars(string value, HtmlRenderBoxStyle? parent, HtmlRenderBoxStyle style) {
        style.HyphenateMinimumWordLength = parent?.HyphenateMinimumWordLength ?? 5;
        style.HyphenateMinimumPrefixLength = parent?.HyphenateMinimumPrefixLength ?? 2;
        style.HyphenateMinimumSuffixLength = parent?.HyphenateMinimumSuffixLength ?? 2;
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return;
        string[] tokens = HtmlRenderCssValues.SplitWhitespace(normalized).ToArray();
        int[] values = tokens
            .Select((token, index) => token == "auto"
                ? index == 0 ? 5 : 2
                : int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 1 ? parsed : -1)
            .ToArray();
        if (values.Length is < 1 or > 3 || values.Any(parsed => parsed < 1)) return;
        style.HyphenateMinimumWordLength = 5;
        style.HyphenateMinimumPrefixLength = 2;
        style.HyphenateMinimumSuffixLength = 2;
        style.HyphenateMinimumWordLength = Math.Min(values[0], 10000);
        if (values.Length >= 2) style.HyphenateMinimumPrefixLength = Math.Min(values[1], 10000);
        if (values.Length >= 3) style.HyphenateMinimumSuffixLength = Math.Min(values[2], 10000);
    }

    private static int? ResolveHyphenateLimitLines(string value, int? inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited;
        if (normalized == "no-limit") return null;
        return int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 1
            ? Math.Min(parsed, 10000)
            : inherited;
    }

    private static string ResolveHyphenateLimitLast(string value, string? inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited ?? "none";
        return normalized == "none" || normalized == "always"
            ? normalized
            : inherited ?? "none";
    }

    private double ResolveHyphenateLimitZone(string value, double containingWidth, double fontSize, double inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited;
        if (TryResolveLength(normalized, containingWidth, fontSize, _options.DefaultFontSize, out double parsed)) return Math.Max(0D, parsed);
        return 0D;
    }

    private void ResolveTabSize(
        string value,
        bool inheritsComputedValue,
        bool resetsToInitial,
        double inherited,
        bool inheritedIsLength,
        double fontSize,
        out double size,
        out bool isLength) {
        if (resetsToInitial) {
            size = 8D;
            isLength = false;
            return;
        }
        string normalized = value.Trim().ToLowerInvariant();
        if (inheritsComputedValue || normalized.Length == 0 || normalized == "inherit" || normalized == "unset") {
            size = inherited;
            isLength = inheritedIsLength;
            return;
        }
        if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            && parsed >= 0D && !double.IsNaN(parsed) && !double.IsInfinity(parsed)
            ) {
            size = parsed;
            isLength = false;
            return;
        }
        if (HtmlRenderCssValues.HasExplicitLengthSyntax(normalized, allowPercentage: false, allowUnitlessZero: true)
            && TryResolveLength(normalized, fontSize, fontSize, _options.DefaultFontSize, out parsed)
            && parsed >= 0D && !double.IsNaN(parsed) && !double.IsInfinity(parsed)) {
            size = parsed;
            isLength = true;
            return;
        }
        size = 8D;
        isLength = false;
    }

    private double ResolveTextSpacing(string value, double fontSize, double inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited;
        if (normalized == "normal") return 0D;
        return TryResolveLength(normalized, fontSize, fontSize, _options.DefaultFontSize, out double parsed)
            && !double.IsNaN(parsed) && !double.IsInfinity(parsed)
            ? parsed
            : 0D;
    }

    private static string ResolveTextOverflow(string value) {
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(value.Trim().ToLowerInvariant());
        return values.Count > 0 && string.Equals(values[values.Count - 1], "ellipsis", StringComparison.Ordinal)
            ? "ellipsis"
            : "clip";
    }

    private static int? ResolveLineClamp(HtmlComputedStyle computed) {
        string value = computed.GetValue("line-clamp").Trim();
        if (value.Length == 0) value = computed.GetValue("-webkit-line-clamp").Trim();
        string token = HtmlRenderCssValues.SplitWhitespace(value).FirstOrDefault() ?? string.Empty;
        return int.TryParse(token, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int parsed)
            && parsed > 0
            ? Math.Min(parsed, 10000)
            : null;
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
                _viewportWidth,
                _viewportHeight,
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
            && !HtmlCssTableParser.TryParseBorderSpacing(borderSpacing, style.Font.Size, _options.DefaultFontSize, _viewportWidth, _viewportHeight, out style.BorderSpacingX, out style.BorderSpacingY)) {
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

    private static string ResolveLanguage(IElement element, string? inherited) {
        string? language = element.GetAttribute("lang");
        if (string.IsNullOrWhiteSpace(language)) language = element.GetAttribute("xml:lang");
        return string.IsNullOrWhiteSpace(language) ? inherited ?? string.Empty : language!.Trim();
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
        if (HtmlRenderCssValues.TryResolveFontSizeKeyword(
                normalized, parentFontSize, _options.DefaultFontSize, out double keywordSize)) {
            return keywordSize;
        }

        return TryResolveLength(normalized, parentFontSize, parentFontSize, _options.DefaultFontSize, out double size) &&
            !double.IsNaN(size) && !double.IsInfinity(size) && size > 0D
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

    private static OfficeTextDecorationStyle ResolveTextDecorationStyle(string value) => value.Trim().ToLowerInvariant() switch {
        "double" => OfficeTextDecorationStyle.Double,
        "dotted" => OfficeTextDecorationStyle.Dotted,
        "dashed" => OfficeTextDecorationStyle.Dashed,
        "wavy" => OfficeTextDecorationStyle.Wavy,
        _ => OfficeTextDecorationStyle.Single
    };

    private static int ResolveTextBaselineLevel(string tag, string value, int inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (tag == "sup" || normalized == "super") return inherited + 1;
        if (tag == "sub" || normalized == "sub") return inherited - 1;
        if (normalized.Length == 0) return inherited;
        if (normalized == "baseline") return inherited;
        return 0;
    }

    private (double Scale, double Offset) ResolveTextBaselineGeometry(
        string tag,
        string value,
        double inheritedScale,
        double inheritedOffset,
        double parentFontSize,
        double parentLineHeight) {
        string normalized = value.Trim().ToLowerInvariant();
        double effectiveParentSize = Math.Max(0.01D, parentFontSize * inheritedScale);
        if (tag == "sup" || normalized == "super") {
            return (inheritedScale * 0.65D, inheritedOffset - effectiveParentSize * 0.30D);
        }
        if (tag == "sub" || normalized == "sub") {
            return (inheritedScale * 0.65D, inheritedOffset + effectiveParentSize * 0.15D);
        }
        if (normalized.Length == 0 || normalized == "baseline") return (inheritedScale, inheritedOffset);
        if (TryResolveLength(normalized, parentLineHeight, effectiveParentSize, _options.DefaultFontSize, out double shift)) {
            return (inheritedScale, inheritedOffset - shift);
        }
        return (inheritedScale, inheritedOffset);
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

    private static string ResolveDisplay(IElement element, string value) {
        if (!string.IsNullOrWhiteSpace(value)) return value.Trim().ToLowerInvariant();
        string tag = element.TagName.ToLowerInvariant();
        if (tag == "math" && string.Equals(element.GetAttribute("display"), "block", StringComparison.OrdinalIgnoreCase)) return "block";
        if (tag == "li") return "list-item";
        if (tag == "table") return "table";
        return IsDefaultBlockTag(tag) ? "block" : "inline";
    }

    private static string ResolvePseudoDisplay(string value) =>
        string.IsNullOrWhiteSpace(value) ? "inline" : value.Trim().ToLowerInvariant();

    private static string ResolveListStyleType(HtmlComputedStyle computed) {
        string type = computed.GetValue("list-style-type").Trim();
        if (type.Length > 0) return type;
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(computed.GetValue("list-style"))) {
            if (string.Equals(token, "none", StringComparison.OrdinalIgnoreCase)) return "none";
            if (!string.Equals(token, "inside", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(token, "outside", StringComparison.OrdinalIgnoreCase)
                && !token.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) return token;
        }

        return string.Empty;
    }

    private static string ResolveListStylePosition(HtmlComputedStyle computed) {
        string position = computed.GetValue("list-style-position").Trim().ToLowerInvariant();
        if (position == "inside" || position == "outside") return position;
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(computed.GetValue("list-style"))) {
            if (string.Equals(token, "inside", StringComparison.OrdinalIgnoreCase)
                || string.Equals(token, "outside", StringComparison.OrdinalIgnoreCase)) return token.ToLowerInvariant();
        }
        return "outside";
    }

    private static string ResolveListStyleImage(HtmlComputedStyle computed) {
        string image = computed.GetValue("list-style-image").Trim();
        if (image.Length > 0) return image;
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(computed.GetValue("list-style"))) {
            if (token.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) return token;
        }
        return "none";
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

    private OfficeColor ResolveColor(IElement element, string value, OfficeColor fallback, bool pseudoElement, string property) {
        string normalized = value.Trim();
        if (normalized.Length == 0 || string.Equals(normalized, "currentcolor", StringComparison.OrdinalIgnoreCase)) return fallback;
        if (HtmlRenderCssValues.TryColor(normalized, out OfficeColor color)) return color;
        ReportUnsupportedColor(element, pseudoElement, property, normalized);
        return fallback;
    }

    private void ReportUnsupportedColor(IElement element, bool pseudoElement, string property, string value) {
        string key = (pseudoElement ? "pseudo:" : string.Empty) + property;
        if (!_reportedUnsupportedColors.TryGetValue(element, out HashSet<string>? properties)) {
            properties = new HashSet<string>(StringComparer.Ordinal);
            _reportedUnsupportedColors[element] = properties;
        }
        if (!properties.Add(key)) return;
        string source = DescribeSource(element) + (pseudoElement ? "::generated" : string.Empty);
        _diagnostics.Add(
            "OfficeIMO.Html.Renderer",
            HtmlRenderDiagnosticCodes.ColorValueUnsupported,
            "A CSS color value outside the static color contract used its property fallback.",
            HtmlDiagnosticSeverity.Warning,
            source,
            property + "=" + value,
            OfficeConversionLossKind.Approximation);
    }

    private static OfficeTextAlignment ResolveAlignment(
        string value,
        string direction,
        OfficeTextAlignment? parentAlignment) {
        if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Center;
        if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Right;
        if (string.Equals(value, "left", StringComparison.OrdinalIgnoreCase)) return OfficeTextAlignment.Left;
        if (string.Equals(value, "match-parent", StringComparison.OrdinalIgnoreCase) && parentAlignment.HasValue) {
            return parentAlignment.Value;
        }
        bool rightToLeft = string.Equals(direction, "rtl", StringComparison.Ordinal);
        if (string.Equals(value, "end", StringComparison.OrdinalIgnoreCase)) return rightToLeft ? OfficeTextAlignment.Left : OfficeTextAlignment.Right;
        return rightToLeft ? OfficeTextAlignment.Right : OfficeTextAlignment.Left;
    }

    private double ResolveLineHeight(string value, double fontSize) {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "normal", StringComparison.OrdinalIgnoreCase)) return fontSize * _options.DefaultLineHeight;
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double multiplier) && multiplier > 0D) return fontSize * multiplier;
        return TryResolveLength(value, fontSize, fontSize, _options.DefaultFontSize, out double lineHeight) && lineHeight > 0D
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

    private static bool PreventsTextWrapping(string tag, string whiteSpace) =>
        tag == "pre" || whiteSpace == "pre" || whiteSpace == "nowrap";

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
        if (margin.Length > 0) HtmlRenderCssValues.ApplyBoxShorthand(
            margin,
            reference,
            fontSize,
            _options.DefaultFontSize,
            _viewportWidth,
            _viewportHeight,
            _activeContainerWidth,
            _activeContainerHeight,
            ref style.MarginTop,
            ref style.MarginRight,
            ref style.MarginBottom,
            ref style.MarginLeft);
        ApplyLength(computed.GetValue("margin-top"), reference, fontSize, ref style.MarginTop);
        ApplyLength(computed.GetValue("margin-right"), reference, fontSize, ref style.MarginRight);
        ApplyLength(computed.GetValue("margin-bottom"), reference, fontSize, ref style.MarginBottom);
        ApplyLength(computed.GetValue("margin-left"), reference, fontSize, ref style.MarginLeft);

        string padding = computed.GetValue("padding");
        if (padding.Length > 0) HtmlRenderCssValues.ApplyBoxShorthand(
            padding,
            reference,
            fontSize,
            _options.DefaultFontSize,
            _viewportWidth,
            _viewportHeight,
            _activeContainerWidth,
            _activeContainerHeight,
            ref style.PaddingTop,
            ref style.PaddingRight,
            ref style.PaddingBottom,
            ref style.PaddingLeft);
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

    private void ApplyPaint(IElement element, HtmlComputedStyle computed, HtmlRenderBoxStyle style, bool pseudoElement) {
        string backgroundShorthand = computed.GetValue("background");
        string background = computed.GetValue("background-color");
        if (background.Length == 0) background = backgroundShorthand;
        if (string.Equals(background.Trim(), "currentcolor", StringComparison.OrdinalIgnoreCase)) {
            style.BackgroundColor = style.Color;
        } else if (HtmlRenderCssValues.TryColor(background, out OfficeColor backgroundColor)) {
            style.BackgroundColor = backgroundColor;
        } else if (computed.GetValue("background-color").Length > 0) {
            ReportUnsupportedColor(element, pseudoElement, "background-color", computed.GetValue("background-color"));
        }
        ApplyBackgroundLayers(computed, style, backgroundShorthand);
        ApplyOpacity(computed.GetValue("opacity"), style);
        style.Transform = NormalizeCssValue(computed.GetValue("transform"), "none");
        style.TransformOrigin = NormalizeCssValue(computed.GetValue("transform-origin"), "50% 50%");
        style.ClipPath = NormalizeCssValue(computed.GetValue("clip-path"), "none");
        style.BoxDecorationBreak = NormalizeCssValue(computed.GetValue("box-decoration-break"), "slice");
        string boxShadow = NormalizeCssValue(computed.GetValue("box-shadow"), "none");
        if (!HtmlCssBoxShadowParser.TryParse(boxShadow, style.Font.Size, _options.DefaultFontSize, _viewportWidth, _viewportHeight, _activeContainerWidth, _activeContainerHeight, style.Color, out IReadOnlyList<HtmlCssBoxShadow> shadows)) {
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
        IReadOnlyList<string> originLayers = HtmlRenderCssValues.SplitTopLevelCommas(computed.GetValue("background-origin"));
        IReadOnlyList<string> clipLayers = HtmlRenderCssValues.SplitTopLevelCommas(computed.GetValue("background-clip"));
        IReadOnlyList<string> attachmentLayers = HtmlRenderCssValues.SplitTopLevelCommas(computed.GetValue("background-attachment"));
        (string shorthandOrigin, string shorthandClip) = ExtractBackgroundBoxes(backgroundShorthand);
        style.BackgroundColorClip = HtmlRenderBackgroundLayer.NormalizeBox(
            clipLayers.Count > 0 ? clipLayers[clipLayers.Count - 1] : shorthandClip,
            "border-box");
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
            string origin = GetLayerValue(originLayers, index, shorthandOrigin, "padding-box");
            string clip = GetLayerValue(clipLayers, index, shorthandClip, "border-box");
            string attachment = GetLayerValue(attachmentLayers, index, ExtractBackgroundAttachment(sourceLayer), "scroll");
            if (urls.Count == 0) {
                if (HtmlCssLinearGradientParser.TryParse(sourceLayer, _options.MaxGradientStops, out HtmlCssLinearGradientDefinition? linearGradient, out bool linearStopLimitExceeded)
                    && linearGradient != null) {
                    layers.Add(new HtmlRenderBackgroundLayer(linearGradient, position, repeat, size, origin, clip, attachment));
                    continue;
                }

                if (HtmlCssRadialGradientParser.TryParse(sourceLayer, _options.MaxGradientStops, out HtmlCssRadialGradientDefinition? radialGradient, out bool radialStopLimitExceeded)
                    && radialGradient != null) {
                    layers.Add(new HtmlRenderBackgroundLayer(radialGradient, position, repeat, size, origin, clip, attachment));
                    continue;
                }

                if (HtmlCssConicGradientParser.TryParse(sourceLayer, _options.MaxGradientStops, out HtmlCssConicGradientDefinition? conicGradient, out bool conicStopLimitExceeded)
                    && conicGradient != null) {
                    layers.Add(new HtmlRenderBackgroundLayer(conicGradient, position, repeat, size, origin, clip, attachment));
                    continue;
                }

                if (linearStopLimitExceeded || radialStopLimitExceeded || conicStopLimitExceeded) {
                    gradientStopLimitExceededCount++;
                } else {
                    unsupportedLayerCount++;
                }

                continue;
            }

            layers.Add(new HtmlRenderBackgroundLayer(urls[0], position, repeat, size, origin, clip, attachment));
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

    private static string ExtractBackgroundAttachment(string shorthand) {
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(shorthand)) {
            string value = token.Trim().TrimEnd(',').ToLowerInvariant();
            if (value == "scroll" || value == "fixed" || value == "local") return value;
        }
        return string.Empty;
    }

    private static (string Origin, string Clip) ExtractBackgroundBoxes(string shorthand) {
        var boxes = new List<string>(2);
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(shorthand)) {
            string value = token.Trim().TrimEnd(',').ToLowerInvariant();
            if (value == "border-box" || value == "padding-box" || value == "content-box") {
                boxes.Add(value);
                if (boxes.Count == 2) break;
            }
        }

        if (boxes.Count == 0) return ("padding-box", "border-box");
        return boxes.Count == 1 ? (boxes[0], boxes[0]) : (boxes[0], boxes[1]);
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
        string position = computed.GetValue("position");
        style.Position = HtmlCssRunningElementParser.TryParsePosition(position, out _)
            ? position.Trim()
            : NormalizeCssValue(position, "static");
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
        return HtmlRenderCssValues.HasExplicitLengthSyntax(normalized, allowPercentage: false, allowUnitlessZero: true)
            && TryResolveLength(normalized, reference, fontSize, _options.DefaultFontSize, out width)
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
        return HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: false, allowUnitlessZero: false)
            && TryResolveLength(value, reference, fontSize, _options.DefaultFontSize, out width)
            && width > 0D;
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
        int index = 0;
        first = ReadPlaceComponent(parts, ref index);
        second = index < parts.Count ? ReadPlaceComponent(parts, ref index) : first;
    }

    private static string ReadPlaceComponent(IReadOnlyList<string> parts, ref int index) {
        string value = parts[index++].Trim().ToLowerInvariant();
        if ((value == "first" || value == "last")
            && index < parts.Count
            && string.Equals(parts[index], "baseline", StringComparison.OrdinalIgnoreCase)) {
            value += " baseline";
            index++;
        }
        return value;
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
        style.RowGapWasSpecified = !string.IsNullOrWhiteSpace(row) && !string.Equals(row.Trim(), "normal", StringComparison.OrdinalIgnoreCase);
        style.RowGap = ResolveGap(row, reference, fontSize, out bool rowUnsupported);
        style.ColumnGap = ResolveGap(column, reference, fontSize, out bool columnUnsupported);
        if (rowUnsupported) style.UnsupportedRowGap = row.Trim();
        if (columnUnsupported) style.UnsupportedColumnGap = column.Trim();
    }

    private double ResolveGap(string value, double reference, double fontSize, out bool unsupported) {
        unsupported = false;
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "normal", StringComparison.OrdinalIgnoreCase)) return 0D;
        if (TryResolveLength(value, reference, fontSize, _options.DefaultFontSize, out double resolved) && resolved >= 0D) return resolved;
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
        if (TryResolveLength(value, reference, fontSize, _options.DefaultFontSize, out double parsed)) target = Math.Max(0D, parsed);
    }

    private double? ReadLength(string cssValue, string? attributeValue, double reference, double fontSize) {
        string value = cssValue.Length > 0 ? cssValue : attributeValue ?? string.Empty;
        return TryResolveLength(value, reference, fontSize, _options.DefaultFontSize, out double parsed) && parsed >= 0D ? parsed : null;
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
