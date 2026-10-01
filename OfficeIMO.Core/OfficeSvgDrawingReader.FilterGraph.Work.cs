using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    // Polygon coverage visits every edge for each of nine samples. Curves flatten
    // to at most 24 segments in the shared renderer; clips incur the same edge work.
    // Charge those costs before synchronous source rasterization, not just nodes.
    private static double EstimateSvgFilterSourceWork(OfficeDrawing drawing, double canvasPixels, CancellationToken token) {
        double work = 0D;
        foreach (OfficeDrawingElement element in drawing.Elements) {
            token.ThrowIfCancellationRequested();
            work += canvasPixels;
            if (element is OfficeDrawingShape shape) {
                work += canvasPixels * 9D * (shape.Shape.Kind switch {
                    OfficeShapeKind.Polygon => shape.Shape.Points.Count,
                    OfficeShapeKind.Path => shape.Shape.PathCommands.Count * 24D,
                    OfficeShapeKind.Ellipse => 72D,
                    OfficeShapeKind.RoundedRectangle => 32D,
                    OfficeShapeKind.Rectangle when shape.Shape.Transform.HasValue => 4D,
                    _ => 0D
                });
                work += canvasPixels * EstimateSvgFilterClipWork(shape.Shape.ClipPath);
            } else if (element is OfficeDrawingGroup group) {
                work += canvasPixels * EstimateSvgFilterClipWork(group.ClipPath) + EstimateSvgFilterSourceWork(group.InnerDrawing, canvasPixels, token);
            } else if (element is OfficeDrawingEffectGroup effect) {
                // Effects render their complete private canvas before compositing,
                // even when the parent filter region or clip is only one pixel.
                double surfacePixels = Math.Ceiling(effect.InnerDrawing.Width) * Math.Ceiling(effect.InnerDrawing.Height);
                work += EstimateSvgFilterSourceWork(effect.InnerDrawing, surfacePixels, token);
                if (effect.SoftMask != null) {
                    OfficeDrawing mask = effect.SoftMask.InnerDrawing;
                    work += surfacePixels + EstimateSvgFilterSourceWork(mask, Math.Ceiling(mask.Width) * Math.Ceiling(mask.Height), token);
                }
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
