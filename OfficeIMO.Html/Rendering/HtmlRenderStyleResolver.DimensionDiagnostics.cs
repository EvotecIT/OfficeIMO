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
            if (property == "width" && style.IntrinsicWidth.HasValue
                || property == "min-width" && style.IntrinsicMinWidth.HasValue
                || property == "max-width" && style.IntrinsicMaxWidth.HasValue) continue;
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

}
