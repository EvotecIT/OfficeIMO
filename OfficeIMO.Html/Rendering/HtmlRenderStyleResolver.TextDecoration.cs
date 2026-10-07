using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<IElement> _reportedDecorationThicknessApproximations = new();

    private void ReportTextDecorationThickness(IElement element, HtmlComputedStyle computed, OfficeFontStyle fontStyle) {
        string thickness = computed.GetValue("text-decoration-thickness");
        if (thickness.Length == 0 || string.Equals(thickness, "auto", StringComparison.OrdinalIgnoreCase)) return;
        bool decorated = (fontStyle & (OfficeFontStyle.Underline | OfficeFontStyle.Strikethrough)) != 0
            || computed.GetValue("text-decoration-line").IndexOf("overline", StringComparison.OrdinalIgnoreCase) >= 0;
        if (!decorated || !_reportedDecorationThicknessApproximations.Add(element)) return;

        _diagnostics.Add(
            "OfficeIMO.Html.Renderer",
            HtmlRenderDiagnosticCodes.TextDecorationThicknessApproximated,
            "Authored text-decoration-thickness uses automatic paint thickness; decoration line, style and color are retained.",
            HtmlDiagnosticSeverity.Warning,
            DescribeSource(element),
            "text-decoration-thickness=" + thickness,
            OfficeConversionLossKind.Approximation);
    }
}
