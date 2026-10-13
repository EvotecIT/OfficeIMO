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

    internal PublisherGroupSpace(PublisherEscherShape group) {
        _coordinates = group.GroupCoordinates;
        _absolute = group.Bounds?.Unrotated(group.Transform.RotationDegrees.GetValueOrDefault());
        Transform = group.GroupTransform;
        ClientTransform = Transform;
        LayoutTransform = group.LayoutTransform;
        ClientLayoutTransform = LayoutTransform;
        Hidden = group.Hidden;
        if (_absolute.HasValue) {
            PublisherNativeRectangle bounds = _absolute.Value;
            var frame = new OfficeImageFrameTransform(group.Transform.RotationDegrees.GetValueOrDefault(),
                (bounds.X1 + bounds.X2) / 2, (bounds.Y1 + bounds.Y2) / 2,
                group.Transform.FlipHorizontal, group.Transform.FlipVertical);
            Transform = frame.CreateDestinationTransform().Then(Transform);
            ClientTransform = Transform;
            if (_coordinates.HasValue) {
                PublisherNativeRectangle coordinates = _coordinates.Value;
                double width = coordinates.X2 - coordinates.X1, height = coordinates.Y2 - coordinates.Y1;
                if (width == 0 || height == 0) throw new InvalidDataException("Publisher group coordinate system is degenerate.");
                Transform = OfficeTransform.Translate(-coordinates.X1, -coordinates.Y1)
                    .Then(OfficeTransform.Scale((bounds.X2 - bounds.X1) / width, (bounds.Y2 - bounds.Y1) / height))
                    .Then(OfficeTransform.Translate(bounds.X1, bounds.Y1)).Then(Transform);
                // Keep layout dimensions in page-sized units without rotating
                // them. The full affine map remains independent for placement.
                LayoutTransform = OfficeTransform.Translate(-Math.Min(coordinates.X1, coordinates.X2), -Math.Min(coordinates.Y1, coordinates.Y2))
                    .Then(OfficeTransform.Scale(Math.Abs((bounds.X2 - bounds.X1) / width), Math.Abs((bounds.Y2 - bounds.Y1) / height)))
                    .Then(OfficeTransform.Translate(Math.Min(bounds.X1, bounds.X2), Math.Min(bounds.Y1, bounds.Y2))).Then(LayoutTransform);
            }
        }
    }

    // Child anchors stay in their parent's coordinate space. Rotate each child
    // there before applying this group's scale, rotation and ancestor maps.
    internal OfficeTransform Transform { get; }
    internal OfficeTransform LayoutTransform { get; }
    // Publisher client anchors already use the enclosing anchor units.
    internal OfficeTransform ClientTransform { get; }
    internal OfficeTransform ClientLayoutTransform { get; }
    internal bool Hidden { get; }

    internal bool TryResolve(PublisherNativeRectangle relative, out PublisherNativeRectangle resolved) {
        resolved = default;
        if (!_coordinates.HasValue || !_absolute.HasValue) return false;
        resolved = relative;
        return true;
    }
}
