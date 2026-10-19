using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public sealed partial class HtmlRenderPage {
    internal HtmlRenderPage CreateRegionInspectionPage(string elementId, CancellationToken cancellationToken) {
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
        return new HtmlRenderPage(PageNumber, region.Width, region.Height,
            content.Select((visual, index) => visual.Translate(-region.X, -region.Y, index)), fonts: _fonts);
    }

    private static IEnumerable<HtmlRenderLayoutRegion> FindRegions(IEnumerable<HtmlRenderVisual> visuals,
        string elementId, CancellationToken cancellationToken) {
        foreach (HtmlRenderVisual visual in visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visual is HtmlRenderLayoutRegion region && region.SourceKey == elementId) yield return region;
            foreach (HtmlRenderLayoutRegion child in FindRegions(InspectionChildren(visual), elementId, cancellationToken)) yield return child;
        }
    }

    private static IEnumerable<HtmlRenderVisual> InspectionChildren(HtmlRenderVisual visual) => visual switch {
        HtmlRenderLayoutRegion r => r.Visuals,
        HtmlRenderSemanticGroup g => g.Visuals,
        HtmlRenderLogicalTextGroup g => g.Visuals,
        HtmlRenderEffectGroup g => g.Visuals,
        HtmlRenderClipGroup g => g.Visuals,
        HtmlRenderPathClipGroup g => g.Visuals,
        HtmlRenderFormField f => f.Visuals,
        _ => Array.Empty<HtmlRenderVisual>()
    };

    // Preserve automatic surface overflow for inspection. Explicit scene clips remain authoritative.
    internal OfficeDrawingQualityReport InspectCanvasBounds(double width, double height, int maximumWidth,
        int maximumHeight, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var content = _scene.Where(visual => visual.Source != "render-surface").ToArray();
        return InspectBounds(content, 0D, 0D, width, height, maximumWidth, maximumHeight, cancellationToken);
    }

    private OfficeDrawingQualityReport InspectBounds(IReadOnlyList<HtmlRenderVisual> content,
        double targetLeft, double targetTop, double width, double height, int maximumWidth,
        int maximumHeight, CancellationToken cancellationToken) {
        // Element containment follows CSS layout rectangles. A positioned text
        // carrier starts one em above its baseline and can extend outside that
        // rectangle without its glyph ink overflowing. Keep ink inspection and
        // actual Drawing/PDF paint on the original scene.
        content = content.Select(ProjectInspectionBounds).ToArray();
        var bounds = ResolveDrawingBufferBounds(content, Math.Max(0D, targetLeft + width), Math.Max(0D, targetTop + height), _fonts, cancellationToken);
        double left = Math.Min(targetLeft, bounds.Left), top = Math.Min(targetTop, bounds.Top);
        double expandedWidth = Math.Max(targetLeft + width, bounds.Right) - left;
        double expandedHeight = Math.Max(targetTop + height, bounds.Bottom) - top;
        if (double.IsNaN(expandedWidth) || double.IsInfinity(expandedWidth) || double.IsNaN(expandedHeight) || double.IsInfinity(expandedHeight) || expandedWidth > maximumWidth || expandedHeight > maximumHeight)
            throw new NotSupportedException("Rendered overflow exceeds the inspection surface limits.");
        var expanded = new HtmlRenderPage(PageNumber, expandedWidth, expandedHeight,
            content.Select((visual, index) => visual.Translate(-left, -top, index)), fonts: _fonts);
        OfficeDrawing drawing = expanded.CreateDrawing(cancellationToken);
        return OfficeDrawingQualityAnalyzer.AnalyzeAtOffset(drawing, targetLeft - left, targetTop - top, width, height,
            new OfficeDrawingQualityOptions(detectTextOverlap: false), cancellationToken);
    }

    private static HtmlRenderVisual ProjectInspectionBounds(HtmlRenderVisual visual) {
        if (visual is HtmlRenderText text && !text.IsPaintOnly) {
            return new HtmlRenderText(text.Text, text.X, text.LayoutY + text.RelativePaintOffsetY,
                text.Width, text.LayoutHeight, text.Font, text.Color, text.Alignment, text.LineHeight,
                text.PaintOrder, text.LinkUri, text.Source, text.SemanticRole,
                layoutY: text.LayoutY, semanticNodeId: text.SemanticNodeId,
                textAdvanceWidth: text.TextAdvanceWidth, layoutHeight: text.LayoutHeight, fontDescriptor: text.FontDescriptor)
                .WithPreparedShaping(text.PreparedShaping);
        }
        return visual.PaintChildren is { } children
            ? visual.ProjectPaintChildren(children.Select(ProjectInspectionBounds), 0D, 0D, visual.PaintOrder)
            : visual;
    }
}
