using OfficeIMO.Drawing;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private static IEnumerable<OfficeDrawingElement> EnumerateDrawingElements(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            yield return element;
            OfficeDrawing? nested = element switch {
                OfficeDrawingGroup group => group.Drawing,
                OfficeDrawingEffectGroup effect => effect.Drawing,
                _ => null
            };
            if (nested == null) continue;
            foreach (OfficeDrawingElement child in EnumerateDrawingElements(nested)) yield return child;
        }
    }
}
