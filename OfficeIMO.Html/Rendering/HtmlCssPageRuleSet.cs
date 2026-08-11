using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssPageRuleSet {
    private readonly List<HtmlCssPageRule> _rules = new List<HtmlCssPageRule>();

    internal void Add(HtmlCssPageRule rule) => _rules.Add(rule);

    internal HtmlCssPageGeometry ResolveGeometry(int pageNumber, string? pageName, HtmlRenderOptions options) {
        IReadOnlyList<HtmlCssPageRule> matching = MatchingRules(pageNumber, pageName).ToList();
        double width = options.PageWidth;
        double height = options.PageHeight;
        string size = matching.Select(rule => rule.Geometry.Size).LastOrDefault(value => value.Length > 0) ?? string.Empty;
        if (size.Length > 0) {
            HtmlCssPageSettingsResolver.TryResolvePageSize(size, options.PageWidth, options.PageHeight, options.DefaultFontSize, out width, out height);
        }

        var resolved = new HtmlCssPageGeometry(width, height, options.Margins);
        foreach (HtmlCssPageRule rule in matching) {
            resolved = rule.Geometry.ApplyMargins(resolved, options);
        }
        return resolved;
    }

    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> ResolveMarginBoxes(int pageNumber, string? pageName) {
        var resolved = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate>();
        foreach (HtmlCssPageRule rule in MatchingRules(pageNumber, pageName)) {
            Apply(rule, resolved);
        }

        return resolved;
    }

    private IEnumerable<HtmlCssPageRule> MatchingRules(int pageNumber, string? pageName) {
        foreach (HtmlCssPageRule rule in _rules.Where(rule => rule.PageName == null && rule.Selector == HtmlCssPageSelector.Generic)) yield return rule;
        foreach (HtmlCssPageRule rule in _rules.Where(rule => rule.PageName == null && rule.Selector != HtmlCssPageSelector.Generic && Matches(rule.Selector, pageNumber))) yield return rule;
        foreach (HtmlCssPageRule rule in _rules.Where(rule => MatchesName(rule.PageName, pageName) && rule.Selector == HtmlCssPageSelector.Generic)) yield return rule;
        foreach (HtmlCssPageRule rule in _rules.Where(rule => MatchesName(rule.PageName, pageName) && rule.Selector != HtmlCssPageSelector.Generic && Matches(rule.Selector, pageNumber))) yield return rule;
    }

    private static void Apply(HtmlCssPageRule rule, IDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> target) {
        foreach (KeyValuePair<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> pair in rule.MarginBoxes) target[pair.Key] = pair.Value;
    }

    private static bool MatchesName(string? ruleName, string? pageName) =>
        ruleName != null && string.Equals(ruleName, pageName, StringComparison.OrdinalIgnoreCase);

    private static bool Matches(HtmlCssPageSelector selector, int pageNumber) {
        if (selector == HtmlCssPageSelector.First) return pageNumber == 1;
        if (selector == HtmlCssPageSelector.Left) return pageNumber % 2 == 0;
        if (selector == HtmlCssPageSelector.Right) return pageNumber % 2 != 0;
        return false;
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

internal readonly struct HtmlCssPageGeometryDeclaration {
    internal HtmlCssPageGeometryDeclaration(
        string size,
        string margin,
        string marginTop,
        string marginRight,
        string marginBottom,
        string marginLeft) {
        Size = size;
        Margin = margin;
        MarginTop = marginTop;
        MarginRight = marginRight;
        MarginBottom = marginBottom;
        MarginLeft = marginLeft;
    }

    internal string Size { get; }
    internal string Margin { get; }
    internal string MarginTop { get; }
    internal string MarginRight { get; }
    internal string MarginBottom { get; }
    internal string MarginLeft { get; }
    internal bool IsEmpty => Size.Length == 0
        && Margin.Length == 0
        && MarginTop.Length == 0
        && MarginRight.Length == 0
        && MarginBottom.Length == 0
        && MarginLeft.Length == 0;

    internal HtmlCssPageGeometry ApplyMargins(HtmlCssPageGeometry current, HtmlRenderOptions options) {
        double width = current.Width;
        double height = current.Height;
        double top = current.Margins.Top;
        double right = current.Margins.Right;
        double bottom = current.Margins.Bottom;
        double left = current.Margins.Left;
        if (Margin.Length > 0) HtmlRenderCssValues.ApplyBoxShorthand(
            Margin,
            width,
            options.DefaultFontSize,
            options.DefaultFontSize,
            width,
            height,
            ref top,
            ref right,
            ref bottom,
            ref left);
        ApplySide(MarginTop, width, height, options.DefaultFontSize, ref top);
        ApplySide(MarginRight, width, height, options.DefaultFontSize, ref right);
        ApplySide(MarginBottom, width, height, options.DefaultFontSize, ref bottom);
        ApplySide(MarginLeft, width, height, options.DefaultFontSize, ref left);
        return new HtmlCssPageGeometry(width, height, new HtmlRenderMargins(left, top, right, bottom));
    }

    private static void ApplySide(string value, double width, double height, double fontSize, ref double target) {
        if (value.Length > 0 && HtmlRenderCssValues.TryLength(value, width, fontSize, fontSize, width, height, out double parsed)) {
            target = Math.Max(0D, parsed);
        }
    }
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
