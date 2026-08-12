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
        return new HtmlCssPageGeometry(width, height, new HtmlRenderMargins(resolvedLeft, resolvedTop, resolvedRight, resolvedBottom));
    }

    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> ResolveMarginBoxes(int pageNumber, string? pageName, HtmlCssPageGeometry geometry, HtmlRenderOptions options) {
        var resolved = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate>();
        foreach (HtmlCssPageRule rule in MatchingRules(pageNumber, pageName)) {
            Apply(rule, resolved, geometry, options);
        }

        return resolved;
    }

    private IEnumerable<HtmlCssPageRule> MatchingRules(int pageNumber, string? pageName) {
        foreach (HtmlCssPageRule rule in _rules.Where(rule => rule.PageName == null && rule.Selector == HtmlCssPageSelector.Generic)) yield return rule;
        foreach (HtmlCssPageRule rule in _rules.Where(rule => rule.PageName == null && rule.Selector != HtmlCssPageSelector.Generic && Matches(rule.Selector, pageNumber))) yield return rule;
        foreach (HtmlCssPageRule rule in _rules.Where(rule => MatchesName(rule.PageName, pageName) && rule.Selector == HtmlCssPageSelector.Generic)) yield return rule;
        foreach (HtmlCssPageRule rule in _rules.Where(rule => MatchesName(rule.PageName, pageName) && rule.Selector != HtmlCssPageSelector.Generic && Matches(rule.Selector, pageNumber))) yield return rule;
    }

    private static void Apply(HtmlCssPageRule rule, IDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> target, HtmlCssPageGeometry geometry, HtmlRenderOptions options) {
        foreach (KeyValuePair<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> pair in rule.MarginBoxes) {
            target[pair.Key] = pair.Value.ResolveViewportUnits(geometry.Width, geometry.Height, options.DefaultFontSize);
        }
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
        string top = parts[0];
        string right = parts.Count > 1 ? parts[1] : top;
        string bottom = parts.Count > 2 ? parts[2] : top;
        string left = parts.Count > 3 ? parts[3] : right;
        sides = new[] { top, right, bottom, left };
        return true;
    }

    private static void ApplySide(HtmlCssPageCascadeValue value, double width, double height, double fontSize, ref double target) {
        if (!value.HasValue) return;
        if (string.Equals(value.Value, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value.Value, "unset", StringComparison.OrdinalIgnoreCase)) {
            target = 0D;
        } else if (HtmlRenderCssValues.TryLength(value.Value, width, fontSize, fontSize, width, height, out double parsed)) {
            target = Math.Max(0D, parsed);
        }
    }

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
}

internal sealed class HtmlCssPageRule {
    internal HtmlCssPageRule(
        string? pageName,
        HtmlCssPageSelector selector,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> marginBoxes,
        HtmlCssPageGeometryDeclaration geometry) {
        PageName = pageName;
        Selector = selector;
        MarginBoxes = marginBoxes;
        Geometry = geometry;
    }

    internal string? PageName { get; }
    internal HtmlCssPageSelector Selector { get; }
    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> MarginBoxes { get; }
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
    internal HtmlCssPageMarginTemplate(HtmlCssPageMarginPosition position, HtmlCssGeneratedContentTemplate content, OfficeFontInfo font, OfficeColor color, OfficeTextAlignment alignment, string fontSizeValue) {
        Position = position;
        Content = content;
        Font = font;
        Color = color;
        Alignment = alignment;
        FontSizeValue = fontSizeValue;
    }

    internal HtmlCssPageMarginPosition Position { get; }
    internal HtmlCssGeneratedContentTemplate Content { get; }
    internal OfficeFontInfo Font { get; }
    internal OfficeColor Color { get; }
    internal OfficeTextAlignment Alignment { get; }
    internal string FontSizeValue { get; }

    internal HtmlCssPageMarginTemplate ResolveViewportUnits(double viewportWidth, double viewportHeight, double defaultFontSize) {
        if (FontSizeValue.Length == 0
            || !HtmlRenderCssValues.TryLength(FontSizeValue, defaultFontSize, defaultFontSize, defaultFontSize, viewportWidth, viewportHeight, out double fontSize)
            || fontSize <= 0D) return this;
        return new HtmlCssPageMarginTemplate(
            Position,
            Content,
            new OfficeFontInfo(Font.FamilyName, fontSize, Font.Style),
            Color,
            Alignment,
            FontSizeValue);
    }
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
