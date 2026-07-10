using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssPageRuleSet {
    private readonly List<HtmlCssPageRule> _rules = new List<HtmlCssPageRule>();

    internal void Add(HtmlCssPageRule rule) => _rules.Add(rule);

    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> ResolveMarginBoxes(int pageNumber, string? pageName) {
        var resolved = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate>();
        foreach (HtmlCssPageRule rule in _rules.Where(rule => rule.PageName == null && rule.Selector == HtmlCssPageSelector.Generic)) {
            Apply(rule, resolved);
        }

        foreach (HtmlCssPageRule rule in _rules.Where(rule => rule.PageName == null && rule.Selector != HtmlCssPageSelector.Generic && Matches(rule.Selector, pageNumber))) {
            Apply(rule, resolved);
        }

        foreach (HtmlCssPageRule rule in _rules.Where(rule => MatchesName(rule.PageName, pageName) && rule.Selector == HtmlCssPageSelector.Generic)) {
            Apply(rule, resolved);
        }

        foreach (HtmlCssPageRule rule in _rules.Where(rule => MatchesName(rule.PageName, pageName) && rule.Selector != HtmlCssPageSelector.Generic && Matches(rule.Selector, pageNumber))) {
            Apply(rule, resolved);
        }

        return resolved;
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
    internal HtmlCssPageRule(string? pageName, HtmlCssPageSelector selector, IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> marginBoxes) {
        PageName = pageName;
        Selector = selector;
        MarginBoxes = marginBoxes;
    }

    internal string? PageName { get; }
    internal HtmlCssPageSelector Selector { get; }
    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> MarginBoxes { get; }
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
