using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssPageRuleSet {
    private readonly List<HtmlCssPageRule> _rules = new List<HtmlCssPageRule>();
    private readonly double? _baseWidth;
    private readonly double? _baseHeight;
    private readonly HtmlRenderMargins? _baseMargins;

    internal HtmlCssPageRuleSet() { }

    internal HtmlCssPageRuleSet(HtmlRenderOptions options) {
        _baseWidth = options.PageWidth;
        _baseHeight = options.PageHeight;
        _baseMargins = options.Margins;
    }

    internal void Add(HtmlCssPageRule rule) {
        rule.SourceOrder = _rules.Count;
        _rules.Add(rule);
    }

    internal HtmlCssPageGeometry ResolveGeometry(int pageNumber, string? pageName, HtmlRenderOptions options) {
        return ResolveGeometry(MatchingRules(pageNumber, pageName), options);
    }

    internal void ApplyGenericGeometry(HtmlRenderOptions options) {
        HtmlCssPageGeometry geometry = ResolveGeometry(
            _rules.Where(rule => rule.PageName == null && rule.Selector == HtmlCssPageSelector.Generic),
            options);
        options.PageSize = new OfficePageSize(
            geometry.Width / HtmlRenderOptions.CssPixelsPerInch,
            geometry.Height / HtmlRenderOptions.CssPixelsPerInch);
        options.Margins = geometry.Margins;
    }

    private HtmlCssPageGeometry ResolveGeometry(IEnumerable<HtmlCssPageRule> rules, HtmlRenderOptions options) {
        IReadOnlyList<HtmlCssPageRule> matching = rules.ToList();
        var size = new HtmlCssPageCascadeValue();
        var top = new HtmlCssPageCascadeValue();
        var right = new HtmlCssPageCascadeValue();
        var bottom = new HtmlCssPageCascadeValue();
        var left = new HtmlCssPageCascadeValue();
        foreach (HtmlCssPageRule rule in matching) {
            HtmlCssPageGeometryDeclaration geometry = rule.Geometry;
            Consider(ref size, geometry.Size, rule);
            if (TryExpandMargin(geometry.Margin.Value, out string[] marginValues)) {
                Consider(ref top, geometry.Margin.WithValue(marginValues[0]), rule);
                Consider(ref right, geometry.Margin.WithValue(marginValues[1]), rule);
                Consider(ref bottom, geometry.Margin.WithValue(marginValues[2]), rule);
                Consider(ref left, geometry.Margin.WithValue(marginValues[3]), rule);
            }
            Consider(ref top, geometry.MarginTop, rule);
            Consider(ref right, geometry.MarginRight, rule);
            Consider(ref bottom, geometry.MarginBottom, rule);
            Consider(ref left, geometry.MarginLeft, rule);
        }

        double width = _baseWidth ?? options.PageWidth;
        double height = _baseHeight ?? options.PageHeight;
        HtmlCssPageCascadeValue effectiveSize = ResolveLayerRevert(size);
        if (effectiveSize.HasValue) {
            HtmlCssPageSettingsResolver.TryResolvePageSize(effectiveSize.Value, width, height, options.DefaultFontSize, out width, out height);
        }

        HtmlRenderMargins baseMargins = _baseMargins ?? options.Margins;
        double resolvedTop = baseMargins.Top;
        double resolvedRight = baseMargins.Right;
        double resolvedBottom = baseMargins.Bottom;
        double resolvedLeft = baseMargins.Left;
        ApplySide(top, width, height, options.DefaultFontSize, ref resolvedTop);
        ApplySide(right, width, height, options.DefaultFontSize, ref resolvedRight);
        ApplySide(bottom, width, height, options.DefaultFontSize, ref resolvedBottom);
        ApplySide(left, width, height, options.DefaultFontSize, ref resolvedLeft);
        return new HtmlCssPageGeometry(width, height, HtmlRenderMargins.FromCssPageRule(resolvedLeft, resolvedTop, resolvedRight, resolvedBottom));
    }

    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> ResolveMarginBoxes(int pageNumber, string? pageName, HtmlCssPageGeometry geometry, HtmlRenderOptions options) {
        IReadOnlyList<HtmlCssPageRule> matching = MatchingRules(pageNumber, pageName).ToList();
        var cascaded = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginCascade>();
        foreach (HtmlCssPageRule rule in matching) {
            foreach (KeyValuePair<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> pair in rule.MarginBoxes) {
                if (!cascaded.TryGetValue(pair.Key, out HtmlCssPageMarginCascade? margin)) {
                    margin = new HtmlCssPageMarginCascade();
                    cascaded[pair.Key] = margin;
                }
                margin.Consider(pair.Value, rule);
            }
        }

        var resolved = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate>();
        foreach (KeyValuePair<HtmlCssPageMarginPosition, HtmlCssPageMarginCascade> pair in cascaded) {
            if (pair.Value.TryBuild(pair.Key, geometry, options, out HtmlCssPageMarginTemplate? template)) {
                resolved[pair.Key] = template!;
            }
        }

        return resolved;
    }

    private IEnumerable<HtmlCssPageRule> MatchingRules(int pageNumber, string? pageName) {
        return _rules.Where(rule => (rule.PageName == null || MatchesName(rule.PageName, pageName))
            && (rule.Selector == HtmlCssPageSelector.Generic || Matches(rule.Selector, pageNumber)));
    }

    private static int PageSelectorSpecificity(HtmlCssPageRule rule) {
        int specificity = rule.PageName == null ? 0 : 4;
        if (rule.Selector == HtmlCssPageSelector.First) specificity += 2;
        else if (rule.Selector == HtmlCssPageSelector.Left || rule.Selector == HtmlCssPageSelector.Right) specificity += 1;
        return specificity;
    }

    private static void Consider(ref HtmlCssPageCascadeValue current, HtmlCssPageDeclaration candidate, HtmlCssPageRule rule) {
        if (candidate.Value.Length == 0) return;
        int specificity = PageSelectorSpecificity(rule);
        var value = new HtmlCssPageCascadeValue(candidate.Value, candidate.IsImportant, rule.LayerOrder, specificity, rule.SourceOrder, candidate.Order);
        if (!current.HasValue) {
            current = value;
        } else if (ShouldReplace(current, candidate, rule.LayerOrder, specificity, rule.SourceOrder)) {
            current = value.WithAlternatives(CollectCandidates(current));
        } else {
            current = current.WithAlternative(value);
        }
    }

    private static bool ShouldReplace(HtmlCssPageCascadeValue current, HtmlCssPageDeclaration candidate, CascadeLayerOrder? layerOrder, int specificity, int sourceOrder) {
        if (candidate.IsImportant != current.IsImportant) return candidate.IsImportant;
        if ((layerOrder != null) != (current.LayerOrder != null)) return candidate.IsImportant ? layerOrder != null : layerOrder == null;
        if (layerOrder != null && current.LayerOrder != null) {
            int layerComparison = layerOrder.CompareTo(current.LayerOrder);
            if (layerComparison != 0) return candidate.IsImportant ? layerComparison < 0 : layerComparison > 0;
        }
        if (specificity != current.Specificity) return specificity > current.Specificity;
        if (sourceOrder != current.SourceOrder) return sourceOrder > current.SourceOrder;
        return candidate.Order >= current.DeclarationOrder;
    }

    internal static bool TryExpandMargin(string value, out string[] sides) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        if (parts.Count < 1 || parts.Count > 4) {
            sides = Array.Empty<string>();
            return false;
        }
        if (parts.Count > 1 && parts.Any(IsCssWidePageMarginKeyword)) {
            sides = Array.Empty<string>();
            return false;
        }
        for (int index = 0; index < parts.Count; index++) {
            if (!IsValidPageMarginComponent(parts[index])) {
                sides = Array.Empty<string>();
                return false;
            }
        }
        string top = parts[0];
        string right = parts.Count > 1 ? parts[1] : top;
        string bottom = parts.Count > 2 ? parts[2] : top;
        string left = parts.Count > 3 ? parts[3] : right;
        sides = new[] { top, right, bottom, left };
        return true;
    }

    private static bool IsCssWidePageMarginKeyword(string value) =>
        string.Equals(value, "initial", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "unset", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "revert-layer", StringComparison.OrdinalIgnoreCase);

    private static void ApplySide(HtmlCssPageCascadeValue value, double width, double height, double fontSize, ref double target) {
        value = ResolveLayerRevert(value);
        if (!value.HasValue) return;
        if (string.Equals(value.Value, "auto", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value.Value, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value.Value, "unset", StringComparison.OrdinalIgnoreCase)) {
            target = 0D;
        } else if (HasPageMarginLengthSyntax(value.Value)
            && HtmlRenderCssValues.TryLength(value.Value, width, fontSize, fontSize, width, height, out double parsed)) {
            target = parsed;
        }
    }

    internal static bool IsValidPageMarginComponent(string value) {
        if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "unset", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "revert-layer", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }
        return HasPageMarginLengthSyntax(value)
            && HtmlRenderCssValues.TryLength(value, 1D, 1D, 1D, 1D, 1D, out _);
    }

    private static bool HasPageMarginLengthSyntax(string value) =>
        HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: true, allowUnitlessZero: true);

    private static bool MatchesName(string? ruleName, string? pageName) =>
        ruleName != null && string.Equals(ruleName, pageName, StringComparison.Ordinal);

    private static bool Matches(HtmlCssPageSelector selector, int pageNumber) {
        if (selector == HtmlCssPageSelector.First) return pageNumber == 1;
        if (selector == HtmlCssPageSelector.Left) return pageNumber % 2 == 0;
        if (selector == HtmlCssPageSelector.Right) return pageNumber % 2 != 0;
        return false;
    }

    private static IReadOnlyList<HtmlCssPageCascadeValue> CollectCandidates(HtmlCssPageCascadeValue value) {
        var candidates = new List<HtmlCssPageCascadeValue>(value.Alternatives.Count + 1) { value };
        candidates.AddRange(value.Alternatives);
        return candidates.AsReadOnly();
    }

    private static HtmlCssPageCascadeValue ResolveLayerRevert(HtmlCssPageCascadeValue value) {
        if (!value.HasValue || !string.Equals(value.Value, "revert-layer", StringComparison.OrdinalIgnoreCase)) return value;
        IReadOnlyList<HtmlCssPageCascadeValue> candidates = CollectCandidates(value);
        var revertedLayers = new HashSet<CascadeLayerOrder?>();
        while (true) {
            HtmlCssPageCascadeValue current = default;
            foreach (HtmlCssPageCascadeValue candidate in candidates) {
                if (revertedLayers.Contains(candidate.LayerOrder)) continue;
                var declaration = new HtmlCssPageDeclaration(candidate.Value, candidate.IsImportant, candidate.DeclarationOrder);
                if (!current.HasValue || ShouldReplace(current, declaration, candidate.LayerOrder, candidate.Specificity, candidate.SourceOrder)) current = candidate;
            }
            if (!current.HasValue || !string.Equals(current.Value, "revert-layer", StringComparison.OrdinalIgnoreCase)) return current;
            if (current.LayerOrder == null) return default;
            revertedLayers.Add(current.LayerOrder);
        }
    }

    private readonly struct HtmlCssPageCascadeValue {
        internal HtmlCssPageCascadeValue(string value, bool isImportant, CascadeLayerOrder? layerOrder, int specificity, int sourceOrder, int declarationOrder, IEnumerable<HtmlCssPageCascadeValue>? alternatives = null) {
            Value = value;
            IsImportant = isImportant;
            LayerOrder = layerOrder;
            Specificity = specificity;
            SourceOrder = sourceOrder;
            DeclarationOrder = declarationOrder;
            HasValue = true;
            Alternatives = new List<HtmlCssPageCascadeValue>(alternatives ?? Array.Empty<HtmlCssPageCascadeValue>()).AsReadOnly();
        }

        internal string Value { get; }
        internal bool IsImportant { get; }
        internal CascadeLayerOrder? LayerOrder { get; }
        internal int Specificity { get; }
        internal int SourceOrder { get; }
        internal int DeclarationOrder { get; }
        internal bool HasValue { get; }
        internal IReadOnlyList<HtmlCssPageCascadeValue> Alternatives { get; }

        internal HtmlCssPageCascadeValue WithAlternative(HtmlCssPageCascadeValue alternative) {
            var alternatives = new List<HtmlCssPageCascadeValue>(Alternatives) { alternative };
            return WithAlternatives(alternatives);
        }

        internal HtmlCssPageCascadeValue WithAlternatives(IEnumerable<HtmlCssPageCascadeValue> alternatives) =>
            new HtmlCssPageCascadeValue(Value, IsImportant, LayerOrder, Specificity, SourceOrder, DeclarationOrder, alternatives);
    }

    private sealed class HtmlCssPageMarginCascade {
        private HtmlCssPageCascadeValue _content;
        private HtmlCssPageCascadeValue _fontFamily;
        private HtmlCssPageCascadeValue _fontSize;
        private HtmlCssPageCascadeValue _fontWeight;
        private HtmlCssPageCascadeValue _fontStyle;
        private HtmlCssPageCascadeValue _color;
        private HtmlCssPageCascadeValue _textAlign;

        internal void Consider(HtmlCssPageMarginRule marginRule, HtmlCssPageRule pageRule) {
            HtmlCssPageRuleSet.Consider(ref _content, marginRule.Content, pageRule);
            HtmlCssPageRuleSet.Consider(ref _fontFamily, marginRule.FontFamily, pageRule);
            HtmlCssPageRuleSet.Consider(ref _fontSize, marginRule.FontSize, pageRule);
            HtmlCssPageRuleSet.Consider(ref _fontWeight, marginRule.FontWeight, pageRule);
            HtmlCssPageRuleSet.Consider(ref _fontStyle, marginRule.FontStyle, pageRule);
            HtmlCssPageRuleSet.Consider(ref _color, marginRule.Color, pageRule);
            HtmlCssPageRuleSet.Consider(ref _textAlign, marginRule.TextAlign, pageRule);
        }

        internal bool TryBuild(HtmlCssPageMarginPosition position, HtmlCssPageGeometry geometry, HtmlRenderOptions options, out HtmlCssPageMarginTemplate? template) {
            template = null;
            if (!_content.HasValue || !HtmlCssGeneratedContentTemplate.TryParse(_content.Value, out HtmlCssGeneratedContentTemplate content)) return false;

            string family = HtmlRenderCssValues.FontFamilyList(_fontFamily.HasValue ? _fontFamily.Value : string.Empty, options.DefaultFontFamily);
            string fontSizeValue = _fontSize.HasValue ? _fontSize.Value : string.Empty;
            double fontSize = options.DefaultFontSize;
            if (!HtmlRenderCssValues.TryLength(fontSizeValue, options.DefaultFontSize, options.DefaultFontSize, options.DefaultFontSize, geometry.Width, geometry.Height, out fontSize)
                || fontSize <= 0D) {
                fontSize = options.DefaultFontSize;
            }

            OfficeFontStyle fontStyle = OfficeFontStyle.Regular;
            string weight = _fontWeight.HasValue ? _fontWeight.Value : string.Empty;
            if (string.Equals(weight, "bold", StringComparison.OrdinalIgnoreCase)
                || int.TryParse(weight, out int numericWeight) && numericWeight >= 600) {
                fontStyle |= OfficeFontStyle.Bold;
            }
            string style = _fontStyle.HasValue ? _fontStyle.Value : string.Empty;
            if (style.StartsWith("italic", StringComparison.OrdinalIgnoreCase)
                || style.StartsWith("oblique", StringComparison.OrdinalIgnoreCase)) {
                fontStyle |= OfficeFontStyle.Italic;
            }

            OfficeColor color = _color.HasValue && HtmlRenderCssValues.TryColor(_color.Value, out OfficeColor parsedColor)
                ? parsedColor
                : OfficeColor.Black;
            OfficeTextAlignment alignment = HtmlCssPageSettingsResolver.ResolveMarginAlignment(
                position,
                _textAlign.HasValue ? _textAlign.Value : string.Empty);
            template = new HtmlCssPageMarginTemplate(
                position,
                content,
                new OfficeFontInfo(family, fontSize, fontStyle),
                color,
                alignment);
            return true;
        }
    }
}

internal sealed class HtmlCssPageRule {
    internal HtmlCssPageRule(
        string? pageName,
        HtmlCssPageSelector selector,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> marginBoxes,
        HtmlCssPageGeometryDeclaration geometry,
        CascadeLayerOrder? layerOrder) {
        PageName = pageName;
        Selector = selector;
        MarginBoxes = marginBoxes;
        Geometry = geometry;
        LayerOrder = layerOrder;
    }

    internal string? PageName { get; }
    internal HtmlCssPageSelector Selector { get; }
    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> MarginBoxes { get; }
    internal HtmlCssPageGeometryDeclaration Geometry { get; }
    internal CascadeLayerOrder? LayerOrder { get; }
    internal int SourceOrder { get; set; }
}

internal readonly struct HtmlCssPageGeometry {
    internal HtmlCssPageGeometry(double width, double height, HtmlRenderMargins margins) {
        Width = width;
        Height = height;
        Margins = margins;
    }

    internal double Width { get; }
    internal double Height { get; }
    internal HtmlRenderMargins Margins { get; }
    internal double ContentWidth => Math.Max(1D, Width - Margins.Left - Margins.Right);
    internal double ContentHeight => Math.Max(1D, Height - Margins.Top - Margins.Bottom);
}

internal sealed class HtmlCssPageMarginRule {
    internal HtmlCssPageMarginRule(
        HtmlCssPageDeclaration content,
        HtmlCssPageDeclaration fontFamily,
        HtmlCssPageDeclaration fontSize,
        HtmlCssPageDeclaration fontWeight,
        HtmlCssPageDeclaration fontStyle,
        HtmlCssPageDeclaration color,
        HtmlCssPageDeclaration textAlign) {
        Content = content;
        FontFamily = fontFamily;
        FontSize = fontSize;
        FontWeight = fontWeight;
        FontStyle = fontStyle;
        Color = color;
        TextAlign = textAlign;
    }

    internal HtmlCssPageDeclaration Content { get; }
    internal HtmlCssPageDeclaration FontFamily { get; }
    internal HtmlCssPageDeclaration FontSize { get; }
    internal HtmlCssPageDeclaration FontWeight { get; }
    internal HtmlCssPageDeclaration FontStyle { get; }
    internal HtmlCssPageDeclaration Color { get; }
    internal HtmlCssPageDeclaration TextAlign { get; }
    internal bool IsEmpty => Content.Value.Length == 0
        && FontFamily.Value.Length == 0
        && FontSize.Value.Length == 0
        && FontWeight.Value.Length == 0
        && FontStyle.Value.Length == 0
        && Color.Value.Length == 0
        && TextAlign.Value.Length == 0;

    internal static HtmlCssPageMarginRule Merge(HtmlCssPageMarginRule earlier, HtmlCssPageMarginRule later) =>
        new HtmlCssPageMarginRule(
            Choose(earlier.Content, later.Content),
            Choose(earlier.FontFamily, later.FontFamily),
            Choose(earlier.FontSize, later.FontSize),
            Choose(earlier.FontWeight, later.FontWeight),
            Choose(earlier.FontStyle, later.FontStyle),
            Choose(earlier.Color, later.Color),
            Choose(earlier.TextAlign, later.TextAlign));

    private static HtmlCssPageDeclaration Choose(HtmlCssPageDeclaration earlier, HtmlCssPageDeclaration later) {
        if (later.Value.Length == 0 || earlier.IsImportant && !later.IsImportant) return earlier;
        return later;
    }
}

internal readonly struct HtmlCssPageDeclaration {
    internal HtmlCssPageDeclaration(string value, bool isImportant, int order) {
        Value = value;
        IsImportant = isImportant;
        Order = order;
    }

    internal string Value { get; }
    internal bool IsImportant { get; }
    internal int Order { get; }
    internal HtmlCssPageDeclaration WithValue(string value) => new HtmlCssPageDeclaration(value, IsImportant, Order);
}

internal readonly struct HtmlCssPageGeometryDeclaration {
    internal HtmlCssPageGeometryDeclaration(
        HtmlCssPageDeclaration size,
        HtmlCssPageDeclaration margin,
        HtmlCssPageDeclaration marginTop,
        HtmlCssPageDeclaration marginRight,
        HtmlCssPageDeclaration marginBottom,
        HtmlCssPageDeclaration marginLeft) {
        Size = size;
        Margin = margin;
        MarginTop = marginTop;
        MarginRight = marginRight;
        MarginBottom = marginBottom;
        MarginLeft = marginLeft;
    }

    internal HtmlCssPageDeclaration Size { get; }
    internal HtmlCssPageDeclaration Margin { get; }
    internal HtmlCssPageDeclaration MarginTop { get; }
    internal HtmlCssPageDeclaration MarginRight { get; }
    internal HtmlCssPageDeclaration MarginBottom { get; }
    internal HtmlCssPageDeclaration MarginLeft { get; }
    internal bool IsEmpty => Size.Value.Length == 0
        && Margin.Value.Length == 0
        && MarginTop.Value.Length == 0
        && MarginRight.Value.Length == 0
        && MarginBottom.Value.Length == 0
        && MarginLeft.Value.Length == 0;
}

internal sealed class HtmlCssPageMarginTemplate {
    internal HtmlCssPageMarginTemplate(HtmlCssPageMarginPosition position, HtmlCssGeneratedContentTemplate content, OfficeFontInfo font, OfficeColor color, OfficeTextAlignment alignment) {
        Position = position;
        Content = content;
        Font = font;
        Color = color;
        Alignment = alignment;
    }

    internal HtmlCssPageMarginPosition Position { get; }
    internal HtmlCssGeneratedContentTemplate Content { get; }
    internal OfficeFontInfo Font { get; }
    internal OfficeColor Color { get; }
    internal OfficeTextAlignment Alignment { get; }
}

internal enum HtmlCssPageSelector {
    Generic,
    First,
    Left,
    Right
}

internal enum HtmlCssPageMarginPosition {
    TopLeftCorner,
    TopLeft,
    TopCenter,
    TopRight,
    TopRightCorner,
    LeftTop,
    LeftMiddle,
    LeftBottom,
    RightTop,
    RightMiddle,
    RightBottom,
    BottomLeftCorner,
    BottomLeft,
    BottomCenter,
    BottomRight,
    BottomRightCorner
}
