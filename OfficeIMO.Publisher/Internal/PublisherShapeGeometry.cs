using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Validates the bounded editable vector contract and resizes its shared geometry without introducing a second shape vocabulary.</summary>
internal static class PublisherShapeGeometry {
    internal static void ValidateBounds(double x, double y, double width, double height, bool line) {
        if (!Finite(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (!Finite(y)) throw new ArgumentOutOfRangeException(nameof(y));
        if (!Finite(width) || (line ? width < 0 : width <= 0)) throw new ArgumentOutOfRangeException(nameof(width));
        if (!Finite(height) || (line ? height < 0 : height <= 0)) throw new ArgumentOutOfRangeException(nameof(height));
        if (line && width == 0 && height == 0) throw new ArgumentException("A line must have two distinct endpoints.");
    }

    internal static void ValidatePaint(double strokeWidth, double fillOpacity, double strokeOpacity) {
        if (!Finite(strokeWidth) || strokeWidth < 0) throw new ArgumentOutOfRangeException(nameof(strokeWidth));
        if (!Finite(fillOpacity) || fillOpacity < 0 || fillOpacity > 1) throw new ArgumentOutOfRangeException(nameof(fillOpacity));
        if (!Finite(strokeOpacity) || strokeOpacity < 0 || strokeOpacity > 1) throw new ArgumentOutOfRangeException(nameof(strokeOpacity));
    }

    internal static OfficeShape OwnedPrimitive(OfficeShape input, double x, double y) {
        OfficeShape shape = input.Clone();
        if (shape.Kind is not (OfficeShapeKind.Rectangle or OfficeShapeKind.RoundedRectangle or OfficeShapeKind.Ellipse or OfficeShapeKind.Line))
            throw new NotSupportedException("Shape creation supports rectangles, rounded rectangles, ellipses and nondegenerate lines.");
        ValidateBounds(x, y, shape.Width, shape.Height, shape.Kind == OfficeShapeKind.Line);
        ValidatePaint(shape.StrokeWidth, shape.FillOpacity ?? 1, shape.StrokeOpacity ?? 1);
        if (shape.Transform.HasValue || shape.ClipPath != null || shape.FillGradient != null || shape.FillRadialGradient != null
            || shape.StrokeGradient != null || shape.StrokeRadialGradient != null || shape.Shadow != null || shape.Glow != null
            || shape.StrokeStartMarker != null || shape.StrokeEndMarker != null || shape.StrokeDashStyle != OfficeStrokeDashStyle.Solid
            || shape.StrokeDashArray.Count != 0 || shape.PathCommands.Count != 0)
            throw new NotSupportedException("Shape creation requires primitive geometry and solid paint without transforms, clipping, effects, dashes or line decorations.");
        if (shape.StrokeLineCap.HasValue && !Enum.IsDefined(typeof(OfficeStrokeLineCap), shape.StrokeLineCap.Value)
            || shape.StrokeLineJoin.HasValue && !Enum.IsDefined(typeof(OfficeStrokeLineJoin), shape.StrokeLineJoin.Value))
            throw new ArgumentException("The shape has an invalid stroke cap or join.", nameof(input));
        if (!Finite(shape.CornerRadius) || shape.CornerRadius < 0
            || shape.Kind == OfficeShapeKind.RoundedRectangle && shape.CornerRadius > Math.Min(shape.Width, shape.Height) / 2)
            throw new ArgumentException("The rounded corner radius must fit the shape.", nameof(input));
        if (shape.Kind == OfficeShapeKind.Line) {
            if (shape.Points.Count != 2 || shape.Points.Any(point => !Finite(point.X) || !Finite(point.Y)
                || point.X < 0 || point.X > shape.Width || point.Y < 0 || point.Y > shape.Height)
                || shape.Points[0] == shape.Points[1]
                || Math.Min(shape.Points[0].X, shape.Points[1].X) != 0 || Math.Max(shape.Points[0].X, shape.Points[1].X) != shape.Width
                || Math.Min(shape.Points[0].Y, shape.Points[1].Y) != 0 || Math.Max(shape.Points[0].Y, shape.Points[1].Y) != shape.Height)
                throw new ArgumentException("A line requires two distinct finite endpoints inside its declared bounds.", nameof(input));
        } else if (shape.Points.Count != 0) throw new ArgumentException("This primitive cannot carry line or polygon points.", nameof(input));
        return shape;
    }

    internal static OfficeShape Resize(OfficeShape original, double width, double height) {
        OfficeShape result;
        if (original.Kind == OfficeShapeKind.Line) {
            // A zero extent is a real horizontal/vertical line, not a tiny diagonal.
            OfficePoint start = original.Points[0], end = original.Points[1];
            result = OfficeShape.Line(start.X <= end.X ? 0 : width, start.Y <= end.Y ? 0 : height,
                start.X <= end.X ? width : 0, start.Y <= end.Y ? height : 0);
        } else if (original.Kind == OfficeShapeKind.Polygon) {
            result = OfficeShape.Polygon(original.Points.Select(point => new OfficePoint(
                point.X * width / original.Width, point.Y * height / original.Height)));
        } else {
            result = original.Clone(); result.Width = width; result.Height = height;
            if (result.Kind == OfficeShapeKind.RoundedRectangle)
                result.CornerRadius = original.CornerRadius * Math.Min(width, height) / Math.Min(original.Width, original.Height);
            return result;
        }
        result.FillColor = original.FillColor; result.StrokeColor = original.StrokeColor;
        result.StrokeWidth = original.StrokeWidth; result.FillOpacity = original.FillOpacity; result.StrokeOpacity = original.StrokeOpacity;
        result.StrokeLineCap = original.StrokeLineCap; result.StrokeLineJoin = original.StrokeLineJoin; result.StrokeMiterLimit = original.StrokeMiterLimit;
        return result;
    }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
