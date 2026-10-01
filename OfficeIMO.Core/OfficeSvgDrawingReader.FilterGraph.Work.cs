using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    // Polygon coverage visits every edge for each of nine samples. Curves flatten
    // to at most 24 segments in the shared renderer; clips incur the same edge work.
    // Charge those costs before synchronous source rasterization, not just nodes.
    private static double EstimateSvgFilterSourceWork(OfficeDrawing drawing, CancellationToken token) {
        double work = 0D;
        foreach (OfficeDrawingElement element in drawing.Elements) {
            token.ThrowIfCancellationRequested();
            work += 1D;
            if (element is OfficeDrawingShape shape) {
                work += 9D * (shape.Shape.Kind switch {
                    OfficeShapeKind.Polygon => shape.Shape.Points.Count,
                    OfficeShapeKind.Path => shape.Shape.PathCommands.Count * 24D,
                    OfficeShapeKind.Ellipse => 72D,
                    OfficeShapeKind.RoundedRectangle => 32D,
                    OfficeShapeKind.Rectangle when shape.Shape.Transform.HasValue => 4D,
                    _ => 0D
                });
                work += EstimateSvgFilterClipWork(shape.Shape.ClipPath);
            } else if (element is OfficeDrawingGroup group) {
                work += EstimateSvgFilterClipWork(group.ClipPath) + EstimateSvgFilterSourceWork(group.InnerDrawing, token);
            } else if (element is OfficeDrawingEffectGroup effect) {
                work += EstimateSvgFilterSourceWork(effect.InnerDrawing, token);
                if (effect.SoftMask != null) work += EstimateSvgFilterSourceWork(effect.SoftMask.InnerDrawing, token);
            }
        }
        return work;
    }

    private static double EstimateSvgFilterClipWork(OfficeClipPath? clip) => clip == null ? 0D : clip.Kind switch {
        OfficeClipPathKind.Path => clip.Commands.Count * 24D,
        OfficeClipPathKind.RoundedRectangle => 32D,
        _ => 0D
    };
}
