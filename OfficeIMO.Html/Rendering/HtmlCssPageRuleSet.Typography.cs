using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Html;

internal sealed partial class HtmlCssPageRuleSet {
    /// <summary>Resolves independent page and margin cascades over the final document-root typography.</summary>
    internal IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> ResolveMarginBoxes(
        int pageNumber, string? pageName, HtmlCssPageGeometry geometry, HtmlRenderOptions options,
        HtmlRenderStyleResolver styleResolver) {
        var resolved = new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate>();
        if (!_rules.Any(rule => rule.MarginBoxes.Count > 0)) return resolved;
        double width = _baseWidth ?? geometry.Width, height = _baseHeight ?? geometry.Height;
        var root = styleResolver.ResolvePageStyleParent();
        var contexts = ResolveMarginContexts(pageNumber, pageName, root.Computed,
            styleResolver.PageComputedStyles, styleResolver.PageCancellationToken);
        if (contexts.Count == 0) return resolved;
        HtmlRenderBoxStyle page = styleResolver.ResolvePageTypography(contexts[0].Page, root.Used, width, height, "@page");
        foreach (var context in contexts) {
            string source = "@page @" + HtmlCssPageSettingsResolver.GetMarginBoxName(context.Position);
            HtmlRenderBoxStyle margin = styleResolver.ResolvePageTypography(context.Margin, page, width, height, source);
            if (!context.Content.IsEmpty) {
                (double containingWidth, double containingHeight) = ResolveMarginContainingDimensions(context.Position, geometry);
                styleResolver.ResolvePageMarginGeometry(context.Margin, margin, width, height,
                    containingWidth, containingHeight, source);
            }
            // The paged-media UA stylesheet sets positional alignment on each
            // margin box. Explicit inherit/unset still selects the page value.
            string alignment = context.Margin.IsImplicitlyInheritedValue("text-align") || context.Margin.IsOriginRevertedValue("text-align")
                ? string.Empty : context.Margin.GetValue("text-align");
            if (context.Margin.IsResetValue("text-align")
                || alignment.Length == 0 && context.Margin.IsInheritedValue("text-align") && !context.Margin.IsImplicitlyInheritedValue("text-align")
                || alignment.Equals("start", StringComparison.OrdinalIgnoreCase)) alignment = "left";
            else if (alignment.Equals("end", StringComparison.OrdinalIgnoreCase)) alignment = "right";
            // Unlike text-align, vertical-align is not implicitly inherited. Its
            // positional UA default applies until this generated context has a value.
            bool hasVerticalAlignment = context.Margin.HasCascadeState("vertical-align")
                && !context.Margin.IsImplicitlyInheritedValue("vertical-align");
            string verticalAlignment = context.Margin.IsOriginRevertedValue("vertical-align")
                ? string.Empty : context.Margin.GetValue("vertical-align");
            if (!context.Margin.IsOriginRevertedValue("vertical-align")
                && (context.Margin.IsResetValue("vertical-align")
                    || verticalAlignment.Length == 0 && context.Margin.IsInheritedValue("vertical-align"))) verticalAlignment = "baseline";
            OfficeTextVerticalAlignment usedVerticalAlignment = HtmlCssPageSettingsResolver.ResolveMarginVerticalAlignment(
                context.Position, verticalAlignment, out string unsupportedVerticalAlignment);
            resolved[context.Position] = new HtmlCssPageMarginTemplate(context.Position, context.Content, margin,
                HtmlCssPageSettingsResolver.ResolveMarginAlignment(context.Position, alignment), usedVerticalAlignment,
                hasVerticalAlignment, unsupportedVerticalAlignment);
        }
        return resolved;
    }

    /// <summary>Supplies percentage references from the physical page-margin rectangle without changing paper units.</summary>
    private static (double Width, double Height) ResolveMarginContainingDimensions(
        HtmlCssPageMarginPosition position, HtmlCssPageGeometry geometry) {
        double contentWidth = Math.Max(0D, geometry.Width - geometry.Margins.Left - geometry.Margins.Right);
        double contentHeight = Math.Max(0D, geometry.Height - geometry.Margins.Top - geometry.Margins.Bottom);
        double left = Math.Max(0D, geometry.Margins.Left), right = Math.Max(0D, geometry.Margins.Right);
        double top = Math.Max(0D, geometry.Margins.Top), bottom = Math.Max(0D, geometry.Margins.Bottom);
        return position switch {
            HtmlCssPageMarginPosition.TopLeft or HtmlCssPageMarginPosition.TopCenter or HtmlCssPageMarginPosition.TopRight => (contentWidth, top),
            HtmlCssPageMarginPosition.BottomLeft or HtmlCssPageMarginPosition.BottomCenter or HtmlCssPageMarginPosition.BottomRight => (contentWidth, bottom),
            HtmlCssPageMarginPosition.LeftTop or HtmlCssPageMarginPosition.LeftMiddle or HtmlCssPageMarginPosition.LeftBottom => (left, contentHeight),
            HtmlCssPageMarginPosition.RightTop or HtmlCssPageMarginPosition.RightMiddle or HtmlCssPageMarginPosition.RightBottom => (right, contentHeight),
            HtmlCssPageMarginPosition.TopLeftCorner => (left, top),
            HtmlCssPageMarginPosition.TopRightCorner => (right, top),
            HtmlCssPageMarginPosition.BottomLeftCorner => (left, bottom),
            HtmlCssPageMarginPosition.BottomRightCorner => (right, bottom),
            _ => throw new ArgumentOutOfRangeException(nameof(position))
        };
    }

    /// <summary>Discovers page-only generic faces through the same cascade used for final painting.</summary>
    internal IEnumerable<(string Families, OfficeFontFaceDescriptor Descriptor)> EnumerateMarginFontRequests(
        HtmlComputedStyle root, HtmlComputedStyleSet styles, HtmlRenderOptions options, CancellationToken cancellationToken,
        bool requireGeneratedContent = false) {
        string rootFamily = HtmlRenderStyleResolver.ResolveFontFamily(string.Empty, root, options.DefaultFontFamily, options.DefaultFontFamily);
        OfficeFontFaceDescriptor rootDescriptor = HtmlRenderStyleResolver.ResolveFontFaceDescriptor(string.Empty, root, OfficeFontFaceDescriptor.Regular);
        foreach (string? name in _rules.Select(rule => rule.PageName).Concat(new string?[] { null }).Distinct(StringComparer.Ordinal)) {
            for (int pageNumber = 1; pageNumber <= 3; pageNumber++) {
                foreach (var context in ResolveMarginContexts(pageNumber, name, root, styles, cancellationToken)) {
                    if (requireGeneratedContent && (context.Content.IsEmpty || context.Content.ContainsRunningElement)) continue;
                    string pageFamily = HtmlRenderStyleResolver.ResolveFontFamily(string.Empty, context.Page, rootFamily, options.DefaultFontFamily);
                    OfficeFontFaceDescriptor pageDescriptor = HtmlRenderStyleResolver.ResolveFontFaceDescriptor(string.Empty, context.Page, rootDescriptor);
                    yield return (HtmlRenderStyleResolver.ResolveFontFamily(string.Empty, context.Margin, pageFamily, options.DefaultFontFamily),
                        HtmlRenderStyleResolver.ResolveFontFaceDescriptor(string.Empty, context.Margin, pageDescriptor));
                }
            }
        }
    }

    /// <summary>Discovers paint resources through the effective generated cascade, including inherited custom properties.</summary>
    internal IEnumerable<HtmlComputedStyle> EnumerateMarginBackgroundStyles(HtmlComputedStyle root,
        HtmlComputedStyleSet styles, CancellationToken cancellationToken) {
        foreach (string? name in _rules.Select(rule => rule.PageName).Concat(new string?[] { null }).Distinct(StringComparer.Ordinal)) {
            for (int pageNumber = 1; pageNumber <= 3; pageNumber++) {
                foreach (var context in ResolveMarginContexts(pageNumber, name, root, styles, cancellationToken)) {
                    if (!context.Content.IsEmpty) yield return context.Margin;
                }
            }
        }
    }

    private IReadOnlyList<MarginStyleContext> ResolveMarginContexts(
        int pageNumber, string? pageName, HtmlComputedStyle root, HtmlComputedStyleSet styles, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlCssPageRule[] rules = MatchingRules(pageNumber, pageName).ToArray();
        if (!rules.Any(rule => rule.MarginBoxes.Count > 0)) return Array.Empty<MarginStyleContext>();
        var key = new MarginStyleContextKey(styles, root, rules);
        if (_marginStyleContexts.TryGetValue(key, out IReadOnlyList<MarginStyleContext>? cached)) return cached;
        var contexts = new List<MarginStyleContext>();
        HtmlComputedStyle page = styles.ComputeGeneratedContext(root,
            rules.SelectMany(rule => MatchDeclarations(rule.StyleDeclarations, rule)), cancellationToken);
        foreach (HtmlCssPageMarginPosition position in rules.SelectMany(rule => rule.MarginBoxes.Keys).Distinct()) {
            cancellationToken.ThrowIfCancellationRequested();
            HtmlCssPageCascadeValue contentValue = default;
            var declarations = new List<HtmlGeneratedStyleDeclaration>();
            foreach (HtmlCssPageRule rule in rules) {
                if (!rule.MarginBoxes.TryGetValue(position, out HtmlCssPageMarginRule? margin)) continue;
                Consider(ref contentValue, margin.Content, rule);
                declarations.AddRange(MatchDeclarations(margin.StyleDeclarations, rule));
            }
            contentValue = ResolveLayerRevert(contentValue);
            if (!contentValue.HasValue || !HtmlCssGeneratedContentTemplate.TryParse(contentValue.Value, out HtmlCssGeneratedContentTemplate content)) continue;
            HtmlComputedStyle marginStyle = styles.ComputeGeneratedContext(page, declarations, cancellationToken);
            contexts.Add(new MarginStyleContext(page, marginStyle, position, content));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return _marginStyleContexts[key] = contexts.ToArray();
    }

    private static IEnumerable<HtmlGeneratedStyleDeclaration> MatchDeclarations(
        IReadOnlyList<HtmlGeneratedStyleDeclaration> declarations, HtmlCssPageRule rule) {
        foreach (HtmlGeneratedStyleDeclaration item in declarations)
            yield return new HtmlGeneratedStyleDeclaration(item.Name, item.Value, item.IsImportant,
                PageSelectorSpecificity(rule), rule.SourceOrder, rule.LayerOrder, item.DeclarationOrder);
    }
}
