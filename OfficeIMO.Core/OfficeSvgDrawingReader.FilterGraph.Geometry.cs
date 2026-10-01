using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    // SVG object bounds exclude filter expansion and clipping. Interactive/paint
    // bounds have a different contract; only the leaf geometry helper is shared.
    private static bool TryGetSvgFilterGeometryBounds(IEnumerable<OfficeDrawingElement> elements, out SvgInteractiveBounds bounds) {
        bounds = default;
        bool found = false;
        foreach (OfficeDrawingElement element in elements) {
            SvgInteractiveBounds current;
            if (element is OfficeDrawingGroup group) {
                if (group.UnfilteredGeometryBounds.HasValue) {
                    var original = group.UnfilteredGeometryBounds.Value;
                    current = new SvgInteractiveBounds(original.Left, original.Top, original.Right, original.Bottom)
                        .Transform(OfficeTransform.Translate(group.X, group.Y));
                } else {
                    if (!TryGetSvgFilterGeometryBounds(group.InnerDrawing.Elements, out current)) continue;
                    current = current.Transform(OfficeTransform.Translate(group.X + group.ContentOffsetX, group.Y + group.ContentOffsetY));
                }
                if (group.FrameTransform.HasValue) current = current.Transform(group.FrameTransform.Value.CreateDestinationTransform());
            } else if (element is OfficeDrawingEffectGroup effect) {
                if (effect.UnfilteredGeometryBounds.HasValue) {
                    var original = effect.UnfilteredGeometryBounds.Value;
                    current = new SvgInteractiveBounds(original.Left, original.Top, original.Right, original.Bottom);
                } else if (!TryGetSvgFilterGeometryBounds(effect.InnerDrawing.Elements, out current)) continue;
                current = current.Transform(effect.Transform);
            } else if (!TryGetSvgElementBounds(element, out current)) {
                continue;
            }
            bounds = found ? bounds.Union(current) : current;
            found = true;
        }
        return found;
    }
}
