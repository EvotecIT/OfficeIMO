using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Maps imported native anchors and group spaces into page-point geometry. Authored picture state never passes through this mapper.</summary>
internal static class PublisherNativeGeometry {
    internal static (double X, double Y, double Width, double Height) Bounds(PublisherEscherShape shape, double originX, double originY) {
        PublisherNativeRectangle native = shape.Bounds!.Value.Unrotated(shape.Transform.RotationDegrees.GetValueOrDefault());
        var mapped = shape.LayoutTransform.TransformRectangleBounds(Math.Min(native.X1, native.X2), Math.Min(native.Y1, native.Y2),
            Math.Abs(native.X2 - native.X1), Math.Abs(native.Y2 - native.Y1));
        return (mapped.Left / 12700 + originX, mapped.Top / 12700 + originY,
            (mapped.Right - mapped.Left) / 12700, (mapped.Bottom - mapped.Top) / 12700);
    }

    internal static OfficeTransform PageTransform(PublisherEscherShape shape, double x, double y, double width, double height, double originX, double originY) {
        if (!shape.LayoutTransform.TryInvert(out _)) return OfficeTransform.Identity;
        double centerX = x + width / 2, centerY = y + height / 2;
        double sx = shape.LayoutTransform.M11, sy = shape.LayoutTransform.M22;
        return OfficeTransform.Translate(-centerX, -centerY).Then(OfficeTransform.Scale(1 / sx, 1 / sy))
            .Then(OfficeTransform.Scale(shape.Transform.FlipHorizontal ? -1 : 1, shape.Transform.FlipVertical ? -1 : 1))
            .Then(OfficeTransform.RotateDegrees(shape.Transform.RotationDegrees.GetValueOrDefault()))
            .Then(OfficeTransform.Scale(sx, sy)).Then(OfficeTransform.Translate(centerX, centerY))
            .Then(GroupPageTransform(shape, originX, originY));
    }

    internal static OfficeTransform GroupPageTransform(PublisherEscherShape shape, double originX, double originY) {
        if (shape.GroupTransform == shape.LayoutTransform) return OfficeTransform.Identity;
        return OfficeTransform.Translate(-originX, -originY).Then(OfficeTransform.Scale(12700, 12700))
            .Then(shape.LayoutTransform.Invert()).Then(shape.GroupTransform).Then(OfficeTransform.Scale(1 / 12700D, 1 / 12700D))
            .Then(OfficeTransform.Translate(originX, originY));
    }
}
