using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private (double Left, double Top, double Right, double Bottom) GroupChildPaintBounds(OfficeDrawing local) {
        double left = 0, top = 0, right = local.Width, bottom = local.Height;
        // Pictures and text retain their nominal frame. Shapes can paint beyond it:
        // custom paths, stroke joins, line markers and shadows all need extra room.
        foreach (OfficeDrawingShape item in local.Elements.OfType<OfficeDrawingShape>()) {
            _context.Token.ThrowIfCancellationRequested();
            OfficeShape shape = item.Shape;
            double padding = 1 + Math.Max(0, shape.StrokeWidth) * Math.Max(1, shape.StrokeMiterLimit) / 2;
            foreach (OfficeLineMarker? marker in new[] { shape.StrokeStartMarker, shape.StrokeEndMarker })
                if (marker != null) padding += Math.Max(marker.Width, marker.Length);
            if (shape.Shadow is OfficeShadow shadow)
                padding += Math.Max(Math.Abs(shadow.OffsetX), Math.Abs(shadow.OffsetY)) + shadow.BlurRadius;
            OfficeTransform placement = (shape.Transform ?? OfficeTransform.Identity)
                .Then(OfficeTransform.Translate(item.X, item.Y));
            foreach (OfficeFlattenedPathContour contour in OfficeStrokeGeometry.FlattenShape(shape, 4))
                foreach (OfficePoint point in contour.Points) {
                    var bounds = placement.TransformRectangleBounds(point.X - padding, point.Y - padding,
                        padding * 2, padding * 2);
                    left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
                    right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
                }
        }
        return (left, top, right, bottom);
    }
}
