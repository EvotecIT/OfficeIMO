using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssPageRuleSet {
    private readonly List<HtmlCssPageRule> _rules = new List<HtmlCssPageRule>();

    internal void Add(HtmlCssPageRule rule) => _rules.Add(rule);

    internal HtmlCssPageGeometry ResolveGeometry(int pageNumber, string? pageName, HtmlRenderOptions options) {
        IReadOnlyList<HtmlCssPageRule> matching = MatchingRules(pageNumber, pageName).ToList();
        var size = new HtmlCssPageCascadeValue();
        var top = new HtmlCssPageCascadeValue();
        var right = new HtmlCssPageCascadeValue();
        var bottom = new HtmlCssPageCascadeValue();
        var left = new HtmlCssPageCascadeValue();
        for (int precedence = 0; precedence < matching.Count; precedence++) {
            HtmlCssPageGeometryDeclaration geometry = matching[precedence].Geometry;
            Consider(ref size, geometry.Size, precedence);
            if (TryExpandMargin(geometry.Margin.Value, out string[] marginValues)) {
                Consider(ref top, geometry.Margin.WithValue(marginValues[0]), precedence);
                Consider(ref right, geometry.Margin.WithValue(marginValues[1]), precedence);
                Consider(ref bottom, geometry.Margin.WithValue(marginValues[2]), precedence);
                Consider(ref left, geometry.Margin.WithValue(marginValues[3]), precedence);
            }
            Consider(ref top, geometry.MarginTop, precedence);
            Consider(ref right, geometry.MarginRight, precedence);
            Consider(ref bottom, geometry.MarginBottom, precedence);
            Consider(ref left, geometry.MarginLeft, precedence);
        }

        double width = options.PageWidth;
        double height = options.PageHeight;
        if (size.HasValue) {
            HtmlCssPageSettingsResolver.TryResolvePageSize(size.Value, options.PageWidth, options.PageHeight, options.DefaultFontSize, out width, out height);
        }

        double resolvedTop = options.Margins.Top;
        double resolvedRight = options.Margins.Right;
        double resolvedBottom = options.Margins.Bottom;
        double resolvedLeft = options.Margins.Left;
        ApplySide(top, width, height, options.DefaultFontSize, ref resolvedTop);
        ApplySide(right, width, height, options.DefaultFontSize, ref resolvedRight);
        ApplySide(bottom, width, height, options.DefaultFontSize, ref resolvedBottom);
        ApplySide(left, width, height, options.DefaultFontSize, ref resolvedLeft);
        return new HtmlCssPageGeometry(width, height, HtmlRenderMargins.FromCssPageRule(resolvedLeft, resolvedTop, resolvedRight, resolvedBottom));
    }

    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> ResolveMarginBoxes(int pageNumber, string? pageName, HtmlCssPageGeometry geometry, HtmlRenderOptions options) {
        IReadOnlyList<HtmlCssPageRule> matching = MatchingRules(pageNumber, pageName).ToList();
        var cascaded = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginCascade>();
        for (int precedence = 0; precedence < matching.Count; precedence++) {
            foreach (KeyValuePair<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> pair in matching[precedence].MarginBoxes) {
                if (!cascaded.TryGetValue(pair.Key, out HtmlCssPageMarginCascade? margin)) {
                    margin = new HtmlCssPageMarginCascade();
                    cascaded[pair.Key] = margin;
                }
                margin.Consider(pair.Value, precedence);
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
        return _rules
            .Where(rule => (rule.PageName == null || MatchesName(rule.PageName, pageName))
                && (rule.Selector == HtmlCssPageSelector.Generic || Matches(rule.Selector, pageNumber)))
            .OrderBy(PageSelectorSpecificity);
    }

    private static int PageSelectorSpecificity(HtmlCssPageRule rule) {
        int specificity = rule.PageName == null ? 0 : 4;
        if (rule.Selector == HtmlCssPageSelector.First) specificity += 2;
        else if (rule.Selector == HtmlCssPageSelector.Left || rule.Selector == HtmlCssPageSelector.Right) specificity += 1;
        return specificity;
    }

    private static void Consider(ref HtmlCssPageCascadeValue current, HtmlCssPageDeclaration candidate, int precedence) {
        if (candidate.Value.Length == 0) return;
        if (!current.HasValue
            || candidate.IsImportant && !current.IsImportant
            || candidate.IsImportant == current.IsImportant && (precedence > current.Precedence
                || precedence == current.Precedence && candidate.Order >= current.DeclarationOrder)) {
            current = new HtmlCssPageCascadeValue(candidate.Value, candidate.IsImportant, precedence, candidate.Order);
        }
    }

    private static bool TryExpandMargin(string value, out string[] sides) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        if (parts.Count < 1 || parts.Count > 4) {
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

    private static void ApplySide(HtmlCssPageCascadeValue value, double width, double height, double fontSize, ref double target) {
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

    private static bool IsValidPageMarginComponent(string value) {
        if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "unset", StringComparison.OrdinalIgnoreCase)) {
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

    private readonly struct HtmlCssPageCascadeValue {
        internal HtmlCssPageCascadeValue(string value, bool isImportant, int precedence, int declarationOrder) {
            Value = value;
            IsImportant = isImportant;
            Precedence = precedence;
            DeclarationOrder = declarationOrder;
            HasValue = true;
        }

        internal string Value { get; }
        internal bool IsImportant { get; }
        internal int Precedence { get; }
        internal int DeclarationOrder { get; }
        internal bool HasValue { get; }
    }

    private sealed class HtmlCssPageMarginCascade {
        private HtmlCssPageCascadeValue _content;
        private HtmlCssPageCascadeValue _fontFamily;
        private HtmlCssPageCascadeValue _fontSize;
        private HtmlCssPageCascadeValue _fontWeight;
        private HtmlCssPageCascadeValue _fontStyle;
        private HtmlCssPageCascadeValue _color;
        private HtmlCssPageCascadeValue _textAlign;

        internal void Consider(HtmlCssPageMarginRule rule, int precedence) {
            HtmlCssPageRuleSet.Consider(ref _content, rule.Content, precedence);
            HtmlCssPageRuleSet.Consider(ref _fontFamily, rule.FontFamily, precedence);
            HtmlCssPageRuleSet.Consider(ref _fontSize, rule.FontSize, precedence);
            HtmlCssPageRuleSet.Consider(ref _fontWeight, rule.FontWeight, precedence);
            HtmlCssPageRuleSet.Consider(ref _fontStyle, rule.FontStyle, precedence);
            HtmlCssPageRuleSet.Consider(ref _color, rule.Color, precedence);
            HtmlCssPageRuleSet.Consider(ref _textAlign, rule.TextAlign, precedence);
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
        HtmlCssPageGeometryDeclaration geometry) {
        PageName = pageName;
        Selector = selector;
        MarginBoxes = marginBoxes;
        Geometry = geometry;
    }

    internal string? PageName { get; }
    internal HtmlCssPageSelector Selector { get; }
    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> MarginBoxes { get; }
    internal HtmlCssPageGeometryDeclaration Geometry { get; }
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
