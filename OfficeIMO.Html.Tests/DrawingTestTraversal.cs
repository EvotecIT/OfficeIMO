using OfficeIMO.Drawing;

namespace OfficeIMO.Tests;

internal static class DrawingTestTraversal {
    internal static IEnumerable<OfficeDrawingElement> Elements(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            yield return element;
            OfficeDrawing? nested = element switch {
                OfficeDrawingGroup group => group.Drawing,
                OfficeDrawingEffectGroup effect => effect.Drawing,
                _ => null
            };
            if (nested == null) continue;
            foreach (OfficeDrawingElement child in Elements(nested)) yield return child;
        }
    }
}
