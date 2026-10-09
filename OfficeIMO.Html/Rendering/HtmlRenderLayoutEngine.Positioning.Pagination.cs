using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Represents a locally positioned box's pure translation as its paged placement.
    /// Its containing block still owns flow height; viewport-fixed and normal-flow paint
    /// effects retain their separate transform/fragmentation behavior.
    /// </summary>
    private bool TryTakePositionedTranslation(
        ref HtmlRenderFlowBlock block,
        HtmlRenderBoxStyle style,
        double containingWidth,
        out double offsetX,
        out double offsetY) {
        offsetX = 0D;
        offsetY = 0D;
        // An unpainted wrapper can expose its child's sole effect group. Require
        // a successfully resolved transform on the positioned box itself;
        // unsupported transforms also leave descendant paint effects untouched.
        if (style.Transform == "none") return false;
        if (block.Visuals.Count != 1 || block.Visuals[0] is not HtmlRenderEffectGroup effect) return false;
        OfficeTransform transform = effect.Transform;
        if (transform.M11 != 1D || transform.M22 != 1D || transform.M12 != 0D || transform.M21 != 0D) return false;
        double boxWidth = ResolveBoxWidth(Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight), style);
        double boxHeight = Math.Max(0.01D, block.Height - style.MarginTop - style.MarginBottom);
        if (!HtmlCssTransformParser.TryParse(
                style.Transform, style.TransformOrigin, style.MarginLeft, style.MarginTop,
                boxWidth, boxHeight, style.Font.Size, _options.DefaultFontSize,
                _activePageGeometry.Width, _activePageGeometry.Height,
                style.ContainerUnitWidth ?? double.NaN, style.ContainerUnitHeight ?? double.NaN,
                out OfficeTransform ownedTransform, out _)
            || !transform.Equals(ownedTransform)) return false;
        offsetX = transform.OffsetX;
        offsetY = transform.OffsetY;
        block = block.WithVisuals(new[] {
            new HtmlRenderEffectGroup(
                effect.X, effect.Y, effect.Width, effect.Height,
                OfficeTransform.Identity, effect.Opacity, effect.Visuals,
                effect.PaintOrder, effect.Source, effect.LayoutY)
        });
        return true;
    }

    /// <summary>
    /// Finds boundaries between non-overlapping positioned boxes in an otherwise
    /// empty paged container. It does not add breaks through normal-flow content
    /// or through overlapping absolute boxes.
    /// </summary>
    private IReadOnlyList<double> CollectPositionedContainerBreakOffsets(
        IElement container,
        double containingWidth,
        double containingHeight,
        double originY,
        double outerHeight) {
        if (_options.Mode != HtmlRenderMode.Paged
            || !_localPositionedElements.TryGetValue(container, out List<PositionedElementRequest>? requests)
            || requests.Count == 0) return Array.Empty<double>();
        var events = new List<(double Offset, bool Starts)>();
        foreach (PositionedElementRequest request in requests) {
            CheckCancellation();
            bool hasRect = _positionedContainingRects.TryGetValue(request.Element, out PositionedContainingRect? rect);
            PositionedLayer layer = request.Resolve(this,
                hasRect ? rect!.Width : containingWidth,
                hasRect ? rect!.Height : containingHeight);
            // A rotation or scale may have a different painted extent. Keep its
            // existing fragmentation behavior rather than inventing row breaks.
            if (!layer.SupportsBoundaryBreaks) return Array.Empty<double>();
            double top = originY + (hasRect ? rect!.Y : 0D) + layer.Y;
            events.Add((top, true));
            events.Add((top + layer.Block.Height, false));
        }
        var offsets = new List<double>();
        int active = 0;
        foreach (IGrouping<double, (double Offset, bool Starts)> boundary in events.GroupBy(item => item.Offset).OrderBy(item => item.Key)) {
            CheckCancellation();
            int ends = boundary.Count(item => !item.Starts);
            if (active - ends == 0 && boundary.Key > 0D && boundary.Key < outerHeight) offsets.Add(boundary.Key);
            active += boundary.Count(item => item.Starts) - ends;
        }
        return offsets;
    }
}
