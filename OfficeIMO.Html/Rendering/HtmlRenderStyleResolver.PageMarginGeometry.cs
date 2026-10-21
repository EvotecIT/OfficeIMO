using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<string> _reportedPageMarginDimensions = new HashSet<string>(StringComparer.Ordinal);

    /// <summary>Resolves generated-box geometry without changing the configured-paper unit or font contexts.</summary>
    internal void ResolvePageMarginGeometry(HtmlComputedStyle computed, HtmlRenderBoxStyle style,
        double paperWidth, double paperHeight, double stripWidth, double stripHeight, string source) {
        if (!computed.HasCascadeStateMatching(HtmlCssPageSettingsResolver.IsMarginGeometryDeclaration)) return;
        double previousWidth = _activeContainerWidth, previousHeight = _activeContainerHeight;
        double previousViewportWidth = _viewportWidth, previousViewportHeight = _viewportHeight;
        double previousAdvance = _activeCharacterAdvance;
        bool previousUpright = _activeUprightVerticalText;
        try {
            _viewportWidth = _activeContainerWidth = paperWidth;
            _viewportHeight = _activeContainerHeight = paperHeight;
            _activeUprightVerticalText = false;
            _activeCharacterAdvance = computed.Properties.Any(pair =>
                (pair.Key is "width" or "min-width" or "max-width" or "padding" || pair.Key.StartsWith("padding-", StringComparison.Ordinal)
                    || pair.Key == "border" || pair.Key.StartsWith("border-", StringComparison.Ordinal)) && ContainsCharacterUnit(pair.Value))
                ? _measureCharacterAdvance(style) : double.NaN;
            style.ExplicitWidth = ReadLength(computed.GetValue("width"), null, stripWidth, style.Font.Size);
            style.MinWidth = ReadLength(computed.GetValue("min-width"), null, stripWidth, style.Font.Size);
            style.MaxWidth = ReadLength(computed.GetValue("max-width"), null, stripWidth, style.Font.Size);
            ApplyLength(computed.GetValue("padding-left"), stripWidth, style.Font.Size, ref style.PaddingLeft);
            ApplyLength(computed.GetValue("padding-right"), stripWidth, style.Font.Size, ref style.PaddingRight);
            // Paged margin contexts use their containing height for vertical percentages.
            ApplyLength(computed.GetValue("padding-top"), stripHeight, style.Font.Size, ref style.PaddingTop);
            ApplyLength(computed.GetValue("padding-bottom"), stripHeight, style.Font.Size, ref style.PaddingBottom);
            if (HtmlCssBoxStrokeParser.TryParseBorder(computed, stripWidth, style.Font.Size, _rootFontSize,
                    paperWidth, paperHeight, paperWidth, paperHeight, style.Color,
                    out HtmlRenderBorderEdges borders, out string detail, _activeCharacterAdvance)) {
                if (Supported(borders.Top) && Supported(borders.Right) && Supported(borders.Bottom) && Supported(borders.Left))
                    style.Borders = borders;
                else style.UnsupportedBorderPaint = "page-margin border-style outside the solid-border profile";
            } else style.UnsupportedBorderPaint = detail;

            foreach (string property in new[] { "width", "min-width", "max-width" }) {
                string value = computed.GetValue(property).Trim();
                if (!(value.Equals("min-content", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("max-content", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("fit-content", StringComparison.OrdinalIgnoreCase)
                    || value.StartsWith("fit-content(", StringComparison.OrdinalIgnoreCase))) continue;
                string identity = source + " " + property + "=" + value;
                if (_reportedPageMarginDimensions.Add(identity)) {
                    _diagnostics.Add("OfficeIMO.Html", HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported,
                        "An intrinsic page-margin width used auto sizing or omitted its minimum or maximum constraint.",
                        HtmlDiagnosticSeverity.Warning, source, property + "=" + value, OfficeConversionLossKind.Approximation);
                }
            }
        } finally {
            _viewportWidth = previousViewportWidth;
            _viewportHeight = previousViewportHeight;
            _activeContainerWidth = previousWidth;
            _activeContainerHeight = previousHeight;
            _activeCharacterAdvance = previousAdvance;
            _activeUprightVerticalText = previousUpright;
        }

        static bool Supported(HtmlRenderBorderSide side) => side.Style is "none" or "hidden" or "solid";
    }
}
