using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<string> _reportedPageMarginDimensions = new HashSet<string>(StringComparer.Ordinal);

    /// <summary>Resolves generated-box geometry without changing the configured-paper unit or font contexts.</summary>
    internal void ResolvePageMarginGeometry(HtmlComputedStyle computed, HtmlRenderBoxStyle style,
        double paperWidth, double paperHeight, double stripWidth, double stripHeight, string source) {
        ResolvePageMarginBackground(computed, style, source);
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
                HtmlCssPageSettingsResolver.IsMarginGeometryDeclaration(pair.Key) && ContainsCharacterUnit(pair.Value))
                ? _measureCharacterAdvance(style) : double.NaN;
            style.BorderBox = string.Equals(computed.GetValue("box-sizing"), "border-box", StringComparison.OrdinalIgnoreCase);
            style.ExplicitWidth = ReadLength(computed.GetValue("width"), null, stripWidth, style.Font.Size);
            style.MinWidth = ReadLength(computed.GetValue("min-width"), null, stripWidth, style.Font.Size);
            style.MaxWidth = ReadLength(computed.GetValue("max-width"), null, stripWidth, style.Font.Size);
            style.IntrinsicWidth = ReadPageMarginIntrinsicWidth("width");
            style.IntrinsicMinWidth = ReadPageMarginIntrinsicWidth("min-width");
            style.IntrinsicMaxWidth = ReadPageMarginIntrinsicWidth("max-width");
            style.ExplicitHeight = ReadPageMarginHeight("height", out style.PageMarginIntrinsicHeight);
            style.MinHeight = ReadPageMarginHeight("min-height", out style.PageMarginIntrinsicMinHeight);
            style.MaxHeight = ReadPageMarginHeight("max-height", out style.PageMarginIntrinsicMaxHeight);
            ApplyAutoMargins(computed, style);
            ApplyMarginLength(computed.GetValue("margin-left"), stripWidth, style.Font.Size, ref style.MarginLeft);
            ApplyMarginLength(computed.GetValue("margin-right"), stripWidth, style.Font.Size, ref style.MarginRight);
            ApplyMarginLength(computed.GetValue("margin-top"), stripHeight, style.Font.Size, ref style.MarginTop);
            ApplyMarginLength(computed.GetValue("margin-bottom"), stripHeight, style.Font.Size, ref style.MarginBottom);
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
        } finally {
            _viewportWidth = previousViewportWidth;
            _viewportHeight = previousViewportHeight;
            _activeContainerWidth = previousWidth;
            _activeContainerHeight = previousHeight;
            _activeCharacterAdvance = previousAdvance;
            _activeUprightVerticalText = previousUpright;
        }

        HtmlRenderIntrinsicWidth? ReadPageMarginIntrinsicWidth(string property) {
            string value = computed.GetValue(property).Trim();
            HtmlRenderIntrinsicWidth? intrinsic = ReadIntrinsicWidth(value, stripWidth, style.Font.Size);
            // Generated margin boxes admit the shared keywords. Functional
            // limits retain their separate, unqualified margin-box contract.
            if (intrinsic.HasValue && !intrinsic.Value.Limit.HasValue) return intrinsic;
            if (!intrinsic.HasValue && !value.StartsWith("fit-content(", StringComparison.OrdinalIgnoreCase)) return null;
            string identity = source + " " + property + "=" + value;
            if (_reportedPageMarginDimensions.Add(identity)) {
                _diagnostics.Add("OfficeIMO.Html", HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported,
                    "An intrinsic page-margin width used auto sizing or omitted its minimum or maximum constraint.",
                    HtmlDiagnosticSeverity.Warning, source, property + "=" + value, OfficeConversionLossKind.Approximation);
            }
            return null;
        }

        double? ReadPageMarginHeight(string property, out bool intrinsic) {
            string value = computed.GetValue(property).Trim();
            intrinsic = value.Equals("min-content", StringComparison.OrdinalIgnoreCase)
                || value.Equals("max-content", StringComparison.OrdinalIgnoreCase)
                || value.Equals("fit-content", StringComparison.OrdinalIgnoreCase);
            if (intrinsic) return null;
            double? height = ReadVerticalLength(value, null, stripHeight, style.Font.Size);
            if (height.HasValue || !value.StartsWith("fit-content(", StringComparison.OrdinalIgnoreCase)) return height;
            if (_reportedPageMarginDimensions.Add(source + " " + property + "=" + value)) {
                _diagnostics.Add("OfficeIMO.Html", HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported,
                    "An intrinsic page-margin height used auto sizing or omitted its minimum or maximum constraint.",
                    HtmlDiagnosticSeverity.Warning, source, property + "=" + value, OfficeConversionLossKind.Approximation);
            }
            return null;
        }

        static bool Supported(HtmlRenderBorderSide side) => side.Style is "none" or "hidden" or "solid";
    }
}
