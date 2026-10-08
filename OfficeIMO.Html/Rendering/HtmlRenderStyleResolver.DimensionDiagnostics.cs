namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<IElement> _reportedIntrinsicDimensions = new HashSet<IElement>();
    private static readonly string[] IntrinsicDimensionProperties = { "width", "min-width", "max-width", "height", "min-height", "max-height" };

    private void ReportUnsupportedIntrinsicDimensions(
        IElement element, HtmlComputedStyle computed, HtmlRenderBoxStyle style, HtmlRenderBoxStyle? parent, bool pseudoElement) {
        if (!AcceptsBoxDimensions(element, style, parent, pseudoElement) || _reportedIntrinsicDimensions.Contains(element)) return;
        List<string>? details = null;
        foreach (string property in IntrinsicDimensionProperties) {
            if (property.EndsWith("width", StringComparison.Ordinal)
                && style.Display is "table-row" or "table-row-group" or "table-header-group" or "table-footer-group") continue;
            if (property.EndsWith("height", StringComparison.Ordinal)
                && style.Display is "table-column" or "table-column-group") continue;
            string value = computed.GetValue(property).Trim();
            if (value.Equals("min-content", StringComparison.OrdinalIgnoreCase)
                || value.Equals("max-content", StringComparison.OrdinalIgnoreCase)
                || value.Equals("fit-content", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("fit-content(", StringComparison.OrdinalIgnoreCase)) {
                (details ??= new List<string>()).Add(property + "=" + value.ToLowerInvariant());
            }
        }
        if (details == null || !_reportedIntrinsicDimensions.Add(element)) return;
        _diagnostics.Add("OfficeIMO.Html", HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported,
            "An intrinsic box-size value used auto sizing or omitted its minimum or maximum constraint.",
            HtmlDiagnosticSeverity.Warning, DescribeSource(element), string.Join(";", details), OfficeConversionLossKind.Approximation);
    }

    private bool AcceptsBoxDimensions(IElement element, HtmlRenderBoxStyle style, HtmlRenderBoxStyle? parent, bool pseudoElement) {
        if (style.Display is "none" or "contents") return false;
        if (style.Display != "inline" || style.Position is "absolute" or "fixed" || style.FloatSide != "none") return true;
        // Floats, positioned boxes and flex/grid items are blockified. Ordinary
        // non-replaced inlines ignore all six dimensions without losing content.
        string parentDisplay = parent?.Display ?? string.Empty;
        // A pseudo's supplied parent style belongs to its originating element.
        IElement? ancestor = pseudoElement ? element : element.ParentElement;
        while (parentDisplay == "contents" && ancestor?.ParentElement != null) {
            ancestor = ancestor.ParentElement;
            parentDisplay = _computedStyles.Elements.TryGetValue(ancestor, out HtmlComputedStyle? ancestorStyle)
                ? ResolveDisplay(ancestor, ancestorStyle.GetValue("display"))
                : HtmlElementDisplay.GetDefaultValue(ancestor);
        }
        if (parentDisplay is "flex" or "inline-flex" or "grid" or "inline-grid") return true;
        if (pseudoElement) return false;
        // Replaced image boxes and native control/math boxes accept dimensions
        // even though their HTML display default is inline in this renderer.
        return element.LocalName.ToLowerInvariant() is "img" or "svg" or "iframe"
            or "input" or "textarea" or "select" or "button" or "math";
    }
}
