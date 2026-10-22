using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlCssPageRuleSet {
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

    internal bool HasPageSpecificRules => _rules.Any(rule => rule.PageName != null || rule.Selector != HtmlCssPageSelector.Generic);

    internal int RuleCount => _rules.Count;

    // Page-context viewport fallback uses the caller's print paper, before @page overrides.
    internal double? InitialPageWidth => _baseWidth;
    internal double? InitialPageHeight => _baseHeight;

    internal HtmlCssPageGeometry ResolveGeometry(int pageNumber, string? pageName, HtmlRenderOptions options) {
        return ResolveGeometry(MatchingRules(pageNumber, pageName), options);
    }

    internal void ApplyGenericGeometry(HtmlRenderOptions options) {
        HtmlCssPageGeometry geometry = ResolveGeometry(
            _rules.Where(rule => rule.PageName == null && rule.Selector == HtmlCssPageSelector.Generic),
            options);
        if (options.PrintFitContentWidth is double layoutWidth) {
            options.CssMediaHeightOverride ??= geometry.Height;
            double physicalContentWidth = geometry.ContentWidth;
            double scale = physicalContentWidth / layoutWidth;
            if (scale >= 1D || layoutWidth > options.MaxSurfaceWidth)
                throw new ArgumentOutOfRangeException(nameof(options.PrintFitContentWidth),
                    "Print layout width must exceed the authored page content width and stay within the surface limit.");
            HtmlCssPageGeometry layoutGeometry = geometry.Scale(1D / scale);
            if (layoutGeometry.Width > options.MaxSurfaceWidth || layoutGeometry.Height > options.MaxSurfaceHeight)
                throw new ArgumentOutOfRangeException(nameof(options.PrintFitContentWidth),
                    "The scaled page exceeds the configured surface limit.");
            options.PrintFitScale = scale;
            geometry = layoutGeometry;
        }
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
        var backgroundColor = new HtmlCssPageCascadeValue();
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
            Consider(ref backgroundColor, geometry.BackgroundColor, rule);
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
        HtmlRenderPrintProductionSettings? printProduction = ResolvePrintProduction(matching, width, height, options);
        HtmlRenderMargins margins = HtmlRenderMargins.FromCssPageRule(resolvedLeft, resolvedTop, resolvedRight, resolvedBottom);
        OfficeColor? resolvedBackgroundColor = ResolveBackgroundColor(backgroundColor);
        if (printProduction == null) return ScaleForPrintFit(
            new HtmlCssPageGeometry(width, height, margins, backgroundColor: resolvedBackgroundColor), options);

        double sheetInset = printProduction.TrimInset;
        return ScaleForPrintFit(new HtmlCssPageGeometry(
            width + (sheetInset * 2D),
            height + (sheetInset * 2D),
            HtmlRenderMargins.FromCssPageRule(
                margins.Left + sheetInset,
                margins.Top + sheetInset,
                margins.Right + sheetInset,
                margins.Bottom + sheetInset),
            printProduction,
            resolvedBackgroundColor), options);
    }

    private static HtmlCssPageGeometry ScaleForPrintFit(HtmlCssPageGeometry geometry, HtmlRenderOptions options) {
        if (options.PrintFitScale is not double scale) return geometry;
        HtmlCssPageGeometry fitted = geometry.Scale(1D / scale);
        if (fitted.Width > options.MaxSurfaceWidth || fitted.Height > options.MaxSurfaceHeight)
            throw new ArgumentOutOfRangeException(nameof(options.PrintFitContentWidth),
                "A fitted CSS page exceeds the configured surface limit.");
        return fitted;
    }

    private static OfficeColor? ResolveBackgroundColor(HtmlCssPageCascadeValue value) {
        value = ResolveLayerRevert(value);
        if (!value.HasValue
            || string.Equals(value.Value, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value.Value, "unset", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value.Value, "transparent", StringComparison.OrdinalIgnoreCase)) return null;
        return HtmlRenderCssValues.TryColor(value.Value, out OfficeColor color) ? color : null;
    }

    private static HtmlRenderPrintProductionSettings? ResolvePrintProduction(
        IReadOnlyList<HtmlCssPageRule> rules,
        double trimWidth,
        double trimHeight,
        HtmlRenderOptions options) {
        var bleed = new HtmlCssPageCascadeValue();
        var marks = new HtmlCssPageCascadeValue();
        foreach (HtmlCssPageRule rule in rules) {
            Consider(ref bleed, rule.Production.Bleed, rule);
            Consider(ref marks, rule.Production.Marks, rule);
        }

        HtmlRenderPrintMarks resolvedMarks = ResolveMarks(ResolveLayerRevert(marks));
        HtmlCssPageCascadeValue resolvedBleedValue = ResolveLayerRevert(bleed);
        double resolvedBleed = 0D;
        if (!resolvedBleedValue.HasValue
            || string.Equals(resolvedBleedValue.Value, "auto", StringComparison.OrdinalIgnoreCase)
            || string.Equals(resolvedBleedValue.Value, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(resolvedBleedValue.Value, "unset", StringComparison.OrdinalIgnoreCase)) {
            if (resolvedMarks != HtmlRenderPrintMarks.None) resolvedBleed = 8D; // CSS Paged Media's 6pt auto bleed.
        } else if (HtmlRenderCssValues.TryLength(
                resolvedBleedValue.Value,
                Math.Min(trimWidth, trimHeight),
                options.DefaultFontSize,
                options.DefaultFontSize,
                trimWidth,
                trimHeight,
                out double parsedBleed)) {
            resolvedBleed = Math.Max(0D, parsedBleed);
        }

        if (resolvedMarks == HtmlRenderPrintMarks.None && resolvedBleed <= 0D) return null;
        double markArea = resolvedMarks == HtmlRenderPrintMarks.None ? 0D : 20D;
        return new HtmlRenderPrintProductionSettings(resolvedBleed, markArea, resolvedMarks);
    }

    private static HtmlRenderPrintMarks ResolveMarks(HtmlCssPageCascadeValue value) {
        if (!value.HasValue
            || string.Equals(value.Value, "none", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value.Value, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value.Value, "unset", StringComparison.OrdinalIgnoreCase)) return HtmlRenderPrintMarks.None;

        HtmlRenderPrintMarks marks = HtmlRenderPrintMarks.None;
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(value.Value)) {
            if (string.Equals(token, "crop", StringComparison.OrdinalIgnoreCase)) marks |= HtmlRenderPrintMarks.Crop;
            else if (string.Equals(token, "cross", StringComparison.OrdinalIgnoreCase)) marks |= HtmlRenderPrintMarks.Cross;
        }
        return marks;
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


}

internal sealed class HtmlCssPageRule {
    internal HtmlCssPageRule(
        string? pageName,
        HtmlCssPageSelector selector,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> marginBoxes,
        HtmlCssPageGeometryDeclaration geometry,
        HtmlCssPageProductionDeclaration production,
        CascadeLayerOrder? layerOrder,
        IReadOnlyList<HtmlGeneratedStyleDeclaration> styleDeclarations) {
        StyleDeclarations = styleDeclarations;
        PageName = pageName;
        Selector = selector;
        MarginBoxes = marginBoxes;
        Geometry = geometry;
        Production = production;
        LayerOrder = layerOrder;
    }

    internal IReadOnlyList<HtmlGeneratedStyleDeclaration> StyleDeclarations { get; }
    internal string? PageName { get; }
    internal HtmlCssPageSelector Selector { get; }
    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> MarginBoxes { get; }
    internal HtmlCssPageGeometryDeclaration Geometry { get; }
    internal HtmlCssPageProductionDeclaration Production { get; }
    internal CascadeLayerOrder? LayerOrder { get; }
    internal int SourceOrder { get; set; }
}

internal readonly struct HtmlCssPageGeometry {
    internal HtmlCssPageGeometry(
        double width,
        double height,
        HtmlRenderMargins margins,
        HtmlRenderPrintProductionSettings? printProduction = null,
        OfficeColor? backgroundColor = null) {
        Width = width;
        Height = height;
        Margins = margins;
        PrintProduction = printProduction;
        BackgroundColor = backgroundColor;
    }

    internal double Width { get; }
    internal double Height { get; }
    internal HtmlRenderMargins Margins { get; }
    internal HtmlRenderPrintProductionSettings? PrintProduction { get; }
    internal OfficeColor? BackgroundColor { get; }
    internal double ContentWidth => Math.Max(1D, Width - Margins.Left - Margins.Right);
    internal double ContentHeight => Math.Max(1D, Height - Margins.Top - Margins.Bottom);

    internal HtmlCssPageGeometry Scale(double factor) => new HtmlCssPageGeometry(
        Width * factor,
        Height * factor,
        HtmlRenderMargins.FromCssPageRule(
            Margins.Left * factor,
            Margins.Top * factor,
            Margins.Right * factor,
            Margins.Bottom * factor),
        PrintProduction == null ? null : new HtmlRenderPrintProductionSettings(
            PrintProduction.Bleed * factor,
            PrintProduction.MarkArea * factor,
            PrintProduction.Marks),
        BackgroundColor);
}

internal readonly struct HtmlCssPageProductionDeclaration {
    internal HtmlCssPageProductionDeclaration(HtmlCssPageDeclaration bleed, HtmlCssPageDeclaration marks) {
        Bleed = bleed;
        Marks = marks;
    }

    internal HtmlCssPageDeclaration Bleed { get; }
    internal HtmlCssPageDeclaration Marks { get; }
    internal bool IsEmpty => Bleed.Value.Length == 0 && Marks.Value.Length == 0;
}

internal sealed class HtmlCssPageMarginRule {
    internal HtmlCssPageMarginRule(HtmlCssPageDeclaration content, IReadOnlyList<HtmlGeneratedStyleDeclaration> styleDeclarations) {
        Content = content;
        StyleDeclarations = styleDeclarations;
    }

    internal HtmlCssPageDeclaration Content { get; }
    internal IReadOnlyList<HtmlGeneratedStyleDeclaration> StyleDeclarations { get; }
    internal bool IsEmpty => Content.Value.Length == 0 && StyleDeclarations.Count == 0;

    internal static HtmlCssPageMarginRule Merge(HtmlCssPageMarginRule earlier, HtmlCssPageMarginRule later) {
        HtmlCssPageDeclaration content = later.Content.Value.Length == 0 || earlier.Content.IsImportant && !later.Content.IsImportant
            ? earlier.Content : later.Content;
        // Repeated boxes in one page rule form one declaration stream. Keep all
        // candidates so layer rollback and variable shorthands retain their owners.
        var declarations = new List<HtmlGeneratedStyleDeclaration>();
        foreach (HtmlGeneratedStyleDeclaration item in earlier.StyleDeclarations.Concat(later.StyleDeclarations)) {
            declarations.Add(new HtmlGeneratedStyleDeclaration(item.Name, item.Value, item.IsImportant,
                0, 0, null, declarations.Count));
        }
        return new HtmlCssPageMarginRule(content, declarations);
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
        HtmlCssPageDeclaration marginLeft,
        HtmlCssPageDeclaration backgroundColor) {
        Size = size;
        Margin = margin;
        MarginTop = marginTop;
        MarginRight = marginRight;
        MarginBottom = marginBottom;
        MarginLeft = marginLeft;
        BackgroundColor = backgroundColor;
    }

    internal HtmlCssPageDeclaration Size { get; }
    internal HtmlCssPageDeclaration Margin { get; }
    internal HtmlCssPageDeclaration MarginTop { get; }
    internal HtmlCssPageDeclaration MarginRight { get; }
    internal HtmlCssPageDeclaration MarginBottom { get; }
    internal HtmlCssPageDeclaration MarginLeft { get; }
    internal HtmlCssPageDeclaration BackgroundColor { get; }
    internal bool IsEmpty => Size.Value.Length == 0
        && Margin.Value.Length == 0
        && MarginTop.Value.Length == 0
        && MarginRight.Value.Length == 0
        && MarginBottom.Value.Length == 0
        && MarginLeft.Value.Length == 0
        && BackgroundColor.Value.Length == 0;
}

internal sealed class HtmlCssPageMarginTemplate {
    internal HtmlCssPageMarginTemplate(HtmlCssPageMarginPosition position, HtmlCssGeneratedContentTemplate content, HtmlRenderBoxStyle typography, OfficeTextAlignment alignment) {
        Position = position;
        Content = content;
        Typography = typography;
        Alignment = alignment;
    }

    internal HtmlCssPageMarginPosition Position { get; }
    internal HtmlCssGeneratedContentTemplate Content { get; }
    internal HtmlRenderBoxStyle Typography { get; }
    internal OfficeFontInfo Font => Typography.Font;
    internal OfficeColor Color => Typography.Color;
    internal OfficeTextAlignment Alignment { get; }
    internal bool HasBoxGeometry => Typography.HasIntrinsicWidths || Typography.ExplicitWidth.HasValue || Typography.MinWidth.HasValue
        || Typography.MaxWidth.HasValue || Typography.HorizontalInsets > 0D || Typography.VerticalInsets > 0D
        || Typography.UnsupportedBorderPaint.Length > 0;
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
