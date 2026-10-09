using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal readonly struct PublisherNativeRectangle {
    internal PublisherNativeRectangle(double x1, double y1, double x2, double y2) {
        if (!Finite(x1) || !Finite(y1) || !Finite(x2) || !Finite(y2))
            throw new InvalidDataException("Publisher group projection produced invalid geometry.");
        X1 = x1; Y1 = y1; X2 = x2; Y2 = y2;
    }
    internal double X1 { get; }
    internal double Y1 { get; }
    internal double X2 { get; }
    internal double Y2 { get; }

    // OfficeArt stores the exchanged anchor dimensions for quarter-turn shapes.
    // Restore the unrotated frame before mapping group child coordinates.
    internal PublisherNativeRectangle Unrotated(double degrees) {
        double angle = ((degrees % 360) + 360) % 360;
        if (angle is not (>= 45 and < 135 or >= 225 and < 315)) return this;
        double centerX = (X1 + X2) / 2, centerY = (Y1 + Y2) / 2;
        double halfWidth = (X2 - X1) / 2, halfHeight = (Y2 - Y1) / 2;
        return new PublisherNativeRectangle(centerX - halfHeight, centerY - halfWidth,
            centerX + halfHeight, centerY + halfWidth);
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

internal readonly struct PublisherGroupSpace {
    private readonly PublisherNativeRectangle? _coordinates;
    private readonly PublisherNativeRectangle? _absolute;

    internal PublisherGroupSpace(PublisherEscherShape group, PublisherGroupSpace? parent) {
        _coordinates = group.GroupCoordinates;
        _absolute = group.Bounds?.Unrotated(group.Transform.RotationDegrees.GetValueOrDefault());
        Transform = parent?.Transform ?? OfficeTransform.Identity;
        Hidden = group.Hidden;
        if (_absolute.HasValue) {
            PublisherNativeRectangle bounds = _absolute.Value;
            var frame = new OfficeImageFrameTransform(group.Transform.RotationDegrees.GetValueOrDefault(),
                (bounds.X1 + bounds.X2) / 2, (bounds.Y1 + bounds.Y2) / 2,
                group.Transform.FlipHorizontal, group.Transform.FlipVertical);
            Transform = frame.CreateDestinationTransform().Then(Transform);
        }
    }

    // Child anchors are mapped into the unrotated, page-centred EMU space.
    // Inherited rotations/reflections remain a separate affine map, so nested
    // groups do not replace rotated rectangles with enlarged bounding boxes.
    internal OfficeTransform Transform { get; }
    internal bool Hidden { get; }

    internal bool TryResolve(PublisherNativeRectangle relative, out PublisherNativeRectangle resolved) {
        resolved = default;
        if (!_coordinates.HasValue || !_absolute.HasValue) return false;
        PublisherNativeRectangle coordinates = _coordinates.Value, absolute = _absolute.Value;
        double width = coordinates.X2 - coordinates.X1, height = coordinates.Y2 - coordinates.Y1;
        if (width == 0 || height == 0) throw new InvalidDataException("Publisher group coordinate system is degenerate.");
        double sx = (absolute.X2 - absolute.X1) / width, sy = (absolute.Y2 - absolute.Y1) / height;
        resolved = new PublisherNativeRectangle(absolute.X1 + (relative.X1 - coordinates.X1) * sx,
            absolute.Y1 + (relative.Y1 - coordinates.Y1) * sy,
            absolute.X1 + (relative.X2 - coordinates.X1) * sx, absolute.Y1 + (relative.Y2 - coordinates.Y1) * sy);
        return true;
    }
}
