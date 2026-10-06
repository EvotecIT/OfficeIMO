using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public sealed partial class HtmlRenderPage {
    internal OfficeDrawingQualityReport InspectRegionBounds(string elementId, int maximumWidth,
        int maximumHeight, CancellationToken cancellationToken) {
        var regions = FindRegions(_scene, elementId, cancellationToken).ToArray();
        if (regions.Length != 1) throw new NotSupportedException("Region inspection requires one rendered positioned, floating, flex or grid region: " + elementId);
        HtmlRenderLayoutRegion region = regions[0];
        // Compare in the region's local layout space. Its own transform moves the box and
        // content together; descendant transforms still affect containment. Ancestor clips
        // do not redefine the region's local layout rectangle.
        var content = region.Visuals.Select(visual => visual is HtmlRenderEffectGroup effect && effect.Source == region.Source
            ? new HtmlRenderEffectGroup(effect.X, effect.Y, effect.Width, effect.Height, OfficeTransform.Identity,
                effect.Opacity, effect.Visuals, effect.PaintOrder, effect.Source, effect.LayoutY)
            : visual);
        var local = new HtmlRenderPage(PageNumber, region.Width, region.Height,
            content.Select((visual, index) => visual.Translate(-region.X, -region.Y, index)), fonts: _fonts);
        return local.InspectCanvasBounds(region.Width, region.Height, maximumWidth, maximumHeight, cancellationToken);
    }

    private static IEnumerable<HtmlRenderLayoutRegion> FindRegions(IEnumerable<HtmlRenderVisual> visuals,
        string elementId, CancellationToken cancellationToken) {
        foreach (HtmlRenderVisual visual in visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visual is HtmlRenderLayoutRegion region && region.SourceKey == elementId) yield return region;
            IEnumerable<HtmlRenderVisual> children = visual switch {
                HtmlRenderLayoutRegion r => r.Visuals,
                HtmlRenderSemanticGroup g => g.Visuals,
                HtmlRenderLogicalTextGroup g => g.Visuals,
                HtmlRenderEffectGroup g => g.Visuals,
                HtmlRenderClipGroup g => g.Visuals,
                HtmlRenderPathClipGroup g => g.Visuals,
                HtmlRenderFormField f => f.Visuals,
                _ => Array.Empty<HtmlRenderVisual>()
            };
            foreach (HtmlRenderLayoutRegion child in FindRegions(children, elementId, cancellationToken)) yield return child;
        }
    }

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
