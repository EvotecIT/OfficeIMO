using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<IElement> _reportedDecorationThicknessApproximations = new();
    private readonly HashSet<IElement> _reportedUnsupportedDecorationLines = new();

    private void ReportTextDecorationLoss(IElement element, HtmlComputedStyle computed, OfficeFontStyle fontStyle) {
        if (HtmlRenderCssValues.SplitWhitespace(computed.GetValue("text-decoration-line"))
            .Contains("overline", StringComparer.OrdinalIgnoreCase)
            && _reportedUnsupportedDecorationLines.Add(element)) {
            _diagnostics.Add(
                "OfficeIMO.Html.Renderer",
                HtmlRenderDiagnosticCodes.TextDecorationLineUnsupported,
                "Authored overline is omitted from static text paint; supported underline and strike-through lines remain available.",
                HtmlDiagnosticSeverity.Warning,
                DescribeSource(element),
                "text-decoration-line=overline",
                OfficeConversionLossKind.Omission);
        }

        string thickness = computed.GetValue("text-decoration-thickness");
        if (thickness.Length == 0 || string.Equals(thickness, "auto", StringComparison.OrdinalIgnoreCase)) return;
        bool decorated = (fontStyle & (OfficeFontStyle.Underline | OfficeFontStyle.Strikethrough)) != 0;
        if (!decorated || !_reportedDecorationThicknessApproximations.Add(element)) return;

        _diagnostics.Add(
            "OfficeIMO.Html.Renderer",
            HtmlRenderDiagnosticCodes.TextDecorationThicknessApproximated,
            "Authored text-decoration-thickness uses automatic paint thickness; supported underline and strike-through style and color are retained.",
            HtmlDiagnosticSeverity.Warning,
            DescribeSource(element),
            "text-decoration-thickness=" + thickness,
            OfficeConversionLossKind.Approximation);
    }
}
