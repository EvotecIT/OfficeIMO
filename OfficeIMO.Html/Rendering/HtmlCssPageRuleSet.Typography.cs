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
        HtmlRenderBoxStyle page = styleResolver.ResolvePageTypography(contexts[0].Page, root.Used, width, height);
        foreach (var context in contexts) {
            HtmlRenderBoxStyle margin = styleResolver.ResolvePageTypography(context.Margin, page, width, height);
            if (!context.Content.IsEmpty && context.Position is (HtmlCssPageMarginPosition.TopLeft or HtmlCssPageMarginPosition.TopCenter or HtmlCssPageMarginPosition.TopRight
                or HtmlCssPageMarginPosition.BottomLeft or HtmlCssPageMarginPosition.BottomCenter or HtmlCssPageMarginPosition.BottomRight)) {
                bool top = context.Position is HtmlCssPageMarginPosition.TopLeft or HtmlCssPageMarginPosition.TopCenter or HtmlCssPageMarginPosition.TopRight;
                styleResolver.ResolvePageMarginGeometry(context.Margin, margin, width, height,
                    Math.Max(0D, geometry.Width - geometry.Margins.Left - geometry.Margins.Right),
                    top ? geometry.Margins.Top : geometry.Margins.Bottom, "@page @" + (context.Position switch {
                        HtmlCssPageMarginPosition.TopLeft => "top-left", HtmlCssPageMarginPosition.TopCenter => "top-center",
                        HtmlCssPageMarginPosition.TopRight => "top-right", HtmlCssPageMarginPosition.BottomLeft => "bottom-left",
                        HtmlCssPageMarginPosition.BottomCenter => "bottom-center", _ => "bottom-right"
                    }));
            }
            // The paged-media UA stylesheet sets positional alignment on each
            // margin box. Explicit inherit/unset still selects the page value.
            string alignment = context.Margin.IsImplicitlyInheritedValue("text-align") || context.Margin.IsOriginRevertedValue("text-align")
                ? string.Empty : context.Margin.GetValue("text-align");
            if (context.Margin.IsResetValue("text-align")
                || alignment.Length == 0 && context.Margin.IsInheritedValue("text-align") && !context.Margin.IsImplicitlyInheritedValue("text-align")
                || alignment.Equals("start", StringComparison.OrdinalIgnoreCase)) alignment = "left";
            else if (alignment.Equals("end", StringComparison.OrdinalIgnoreCase)) alignment = "right";
            resolved[context.Position] = new HtmlCssPageMarginTemplate(context.Position, context.Content, margin,
                HtmlCssPageSettingsResolver.ResolveMarginAlignment(context.Position, alignment));
        }
        return resolved;
    }

    /// <summary>Discovers page-only generic faces through the same cascade used for final painting.</summary>
    internal IEnumerable<(string Families, OfficeFontFaceDescriptor Descriptor)> EnumerateMarginFontRequests(
        HtmlComputedStyle root, HtmlComputedStyleSet styles, HtmlRenderOptions options, CancellationToken cancellationToken) {
        string rootFamily = HtmlRenderStyleResolver.ResolveFontFamily(string.Empty, root, options.DefaultFontFamily, options.DefaultFontFamily);
        OfficeFontFaceDescriptor rootDescriptor = HtmlRenderStyleResolver.ResolveFontFaceDescriptor(string.Empty, root, OfficeFontFaceDescriptor.Regular);
        foreach (string? name in _rules.Select(rule => rule.PageName).Concat(new string?[] { null }).Distinct(StringComparer.Ordinal)) {
            for (int pageNumber = 1; pageNumber <= 3; pageNumber++) {
                foreach (var context in ResolveMarginContexts(pageNumber, name, root, styles, cancellationToken)) {
                    string pageFamily = HtmlRenderStyleResolver.ResolveFontFamily(string.Empty, context.Page, rootFamily, options.DefaultFontFamily);
                    OfficeFontFaceDescriptor pageDescriptor = HtmlRenderStyleResolver.ResolveFontFaceDescriptor(string.Empty, context.Page, rootDescriptor);
                    yield return (HtmlRenderStyleResolver.ResolveFontFamily(string.Empty, context.Margin, pageFamily, options.DefaultFontFamily),
                        HtmlRenderStyleResolver.ResolveFontFaceDescriptor(string.Empty, context.Margin, pageDescriptor));
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
