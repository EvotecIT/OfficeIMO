using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static OfficeLinearGradient TransformShapeFillGradient(OfficeDrawingShape shape, double scale,
        IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeLinearGradient gradient) =>
        gradient.TransformCoordinates(ShapeFillCoordinates(shape, scale, contours));

    private static OfficeRadialGradient TransformShapeFillGradient(OfficeDrawingShape shape, double scale,
        IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeRadialGradient gradient) =>
        gradient.TransformCoordinates(ShapeFillCoordinates(shape, scale, contours));

    private static OfficeTransform ShapeFillCoordinates(OfficeDrawingShape drawingShape, double scale,
        IReadOnlyList<IReadOnlyList<OfficePoint>> transformedContours) {
        OfficeShape shape = drawingShape.Shape;
        var coordinates = OfficeTransform.Scale(shape.Width, shape.Height)
            .Then(shape.Transform ?? OfficeTransform.Identity)
            .Then(OfficeTransform.Translate(drawingShape.X, drawingShape.Y))
            .Then(OfficeTransform.Scale(scale, scale));
        return NormalizePaintCoordinates(coordinates, transformedContours);
    }

    private static OfficeTransform NormalizePaintCoordinates(OfficeTransform coordinates,
        IReadOnlyList<IReadOnlyList<OfficePoint>> contours) {
        if (!TryGetContourBounds(contours, out double left, out double top,
            out double right, out double bottom)) return OfficeTransform.Identity;
        return coordinates.Then(OfficeTransform.Translate(-left, -top))
            .Then(OfficeTransform.Scale(1D / (right - left), 1D / (bottom - top)));
    }
}
