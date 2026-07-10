using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool HasNonVisibleOverflow(HtmlRenderBoxStyle style) =>
        style.OverflowX != "visible" || style.OverflowY != "visible";

    private void ApplyViewportOverflow(IList<HtmlRenderVisual> visuals, double width, double height) {
        if (_viewportOverflowElement == null || _viewportOverflowStyle == null || !HasNonVisibleOverflow(_viewportOverflowStyle)) return;
        var content = visuals
            .Where(visual => !IsSurfacePaint(visual))
            .OrderBy(visual => visual.PaintOrder)
            .Select((visual, index) => visual.Translate(0D, 0D, index))
            .ToList();
        if (content.Count == 0) return;

        for (int index = visuals.Count - 1; index >= 0; index--) {
            if (!IsSurfacePaint(visuals[index])) visuals.RemoveAt(index);
        }

        bool clipHorizontal = _viewportOverflowStyle.OverflowX != "visible";
        bool clipVertical = _viewportOverflowStyle.OverflowY != "visible";
        if (ShouldReportOverflowSnapshot(content, _viewportOverflowStyle, 0D, 0D, width, height)
            && _reportedOverflowScrollSnapshots.Add(_viewportOverflowElement)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.OverflowScrollSnapshot,
                "The propagated root overflow rendered the initial static viewport without interactive scrollbars.",
                HtmlDiagnosticSeverity.Info,
                HtmlRenderStyleResolver.DescribeSource(_viewportOverflowElement),
                "overflow-x=" + _viewportOverflowStyle.OverflowX + ";overflow-y=" + _viewportOverflowStyle.OverflowY);
        }

        visuals.Add(new HtmlRenderClipGroup(
            0D,
            0D,
            width,
            height,
            clipHorizontal,
            clipVertical,
            content,
            visuals.Count,
            HtmlRenderStyleResolver.DescribeSource(_viewportOverflowElement) + ":viewport-overflow"));
    }

    private static bool IsSurfacePaint(HtmlRenderVisual visual) =>
        string.Equals(visual.Source, "render-surface", StringComparison.Ordinal)
        || visual.Source != null && visual.Source.StartsWith("render-root-background", StringComparison.Ordinal);
}
