using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static class HtmlCssPrintFitResolver {
    internal static bool TryApplyWideRoot(
        IHtmlDocument document,
        HtmlComputedStyleSet styles,
        HtmlCssPageRuleSet pageRules,
        HtmlRenderOptions options) {
        if (!options.AutoFitWidePrintRoot || options.Mode != HtmlRenderMode.Paged
            || !options.HonorCssPageRules || options.PrintFitContentWidth.HasValue
            || pageRules.HasPageSpecificRules) return false;

        double contentWidth = options.PageWidth - options.Margins.Left - options.Margins.Right;
        if (contentWidth <= 0D) return false;
        var resolver = new HtmlRenderStyleResolver(styles, options, new HtmlDiagnosticReport());
        IElement? root = document.DocumentElement;
        if (root == null) return false;
        HtmlRenderBoxStyle rootStyle = resolver.Resolve(root, contentWidth);
        double width = FixedMinimumWidth(root, rootStyle, styles);
        if (document.Body is IElement body) {
            HtmlRenderBoxStyle bodyStyle = resolver.Resolve(body, contentWidth, rootStyle);
            width = Math.Max(width, FixedMinimumWidth(body, bodyStyle, styles));
        }
        if (width <= contentWidth + 0.5D || width > options.MaxSurfaceWidth) return false;

        double expansion = width / contentWidth;
        if (options.PageWidth * expansion > options.MaxSurfaceWidth
            || options.PageHeight * expansion > options.MaxSurfaceHeight) return false;

        // CSS and resource selection already used the physical print media width.
        // Keep that media context while the layout surface expands for PDF fitting.
        options.CssMediaWidthOverride = options.CssMediaWidth;
        options.CssMediaHeightOverride = options.CssMediaHeight;
        options.PrintFitContentWidth = width;
        pageRules.ApplyGenericGeometry(options);
        return true;
    }

    private static double FixedMinimumWidth(IElement element, HtmlRenderBoxStyle box, HtmlComputedStyleSet styles) {
        if (!styles.Elements.TryGetValue(element, out HtmlComputedStyle? style)
            || !HasFixedLengthUnit(style.GetValue("min-width"))) return 0D;
        return box.MinWidth.GetValueOrDefault()
            + (box.BorderBox ? 0D : box.HorizontalInsets)
            + Math.Max(0D, box.MarginLeft + box.MarginRight);
    }

    private static bool HasFixedLengthUnit(string value) {
        string trimmed = value.Trim();
        return trimmed.EndsWith("px", StringComparison.OrdinalIgnoreCase)
            || trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase)
            || trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase)
            || trimmed.EndsWith("mm", StringComparison.OrdinalIgnoreCase)
            || trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
            || trimmed.EndsWith("pc", StringComparison.OrdinalIgnoreCase)
            || trimmed.EndsWith("q", StringComparison.OrdinalIgnoreCase);
    }
}
