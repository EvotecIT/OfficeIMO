using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public sealed partial class HtmlRenderPage {
    // Preserve automatic surface overflow for inspection. Explicit scene clips remain authoritative.
    internal OfficeDrawingQualityReport InspectCanvasBounds(double width, double height, int maximumWidth,
        int maximumHeight, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var content = _scene.Where(visual => visual.Source != "render-surface").ToArray();
        var bounds = ResolveDrawingBufferBounds(content, width, height, _fonts, cancellationToken);
        double left = Math.Min(0D, bounds.Left), top = Math.Min(0D, bounds.Top);
        double expandedWidth = Math.Max(width, bounds.Right) - left;
        double expandedHeight = Math.Max(height, bounds.Bottom) - top;
        if (double.IsNaN(expandedWidth) || double.IsInfinity(expandedWidth) || double.IsNaN(expandedHeight) || double.IsInfinity(expandedHeight) || expandedWidth > maximumWidth || expandedHeight > maximumHeight)
            throw new NotSupportedException("Rendered overflow exceeds the inspection surface limits.");
        var expanded = new HtmlRenderPage(PageNumber, expandedWidth, expandedHeight,
            content.Select((visual, index) => visual.Translate(-left, -top, index)), fonts: _fonts);
        OfficeDrawing drawing = expanded.CreateDrawing(cancellationToken);
        return OfficeDrawingQualityAnalyzer.AnalyzeAtOffset(drawing, -left, -top, width, height,
            new OfficeDrawingQualityOptions(detectTextOverlap: false), cancellationToken);
    }
}
