using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    /// <summary>Maps every object through its owning page's native anchor origin.</summary>
    private (double X, double Y) PageOrigin(PublisherEscherShape shape) {
        uint? owner = Owner(shape.Id, 0);
        return owner.HasValue && _pages.TryGetValue(owner.Value, out PublisherSourcePage? page)
            ? PageOrigin(page) : (_source.Width / 2, _source.Height / 2);
    }

    private (double X, double Y) PageOrigin(PublisherSourcePage page) {
        double x = _source.Width / 2;
        if (page.IsMaster) {
            // Paired master layout codes select the left or right page edge
            // as the horizontal anchor origin instead of the page center.
            // Single masters and physical pages use the page-centered basis.
            x = page.NativeLayout switch { 0 => 0, 1 => _source.Width, _ => x };
        }
        return (x, _source.Height / 2);
    }

    private FrameRectangle FrameBounds(PublisherEscherShape shape) {
        PublisherNativeRectangle native = shape.Bounds!.Value.Unrotated(shape.Transform.RotationDegrees.GetValueOrDefault());
        var mapped = shape.LayoutTransform.TransformRectangleBounds(Math.Min(native.X1, native.X2), Math.Min(native.Y1, native.Y2),
            Math.Abs(native.X2 - native.X1), Math.Abs(native.Y2 - native.Y1));
        var origin = PageOrigin(shape);
        double x = mapped.Left / 12700 + origin.X, y = mapped.Top / 12700 + origin.Y;
        double width = (mapped.Right - mapped.Left) / 12700, height = (mapped.Bottom - mapped.Top) / 12700;
        return new FrameRectangle(x, y, width, height);
    }

    private OfficeTransform PageTransform(PublisherEscherShape shape, FrameRectangle bounds) {
        // Collapsed layouts have no inverse frame basis. Their artwork is
        // omitted by the geometry guard; retained model positions stay finite.
        if (!shape.LayoutTransform.TryInvert(out _)) return OfficeTransform.Identity;
        double x = bounds.X + bounds.Width / 2, y = bounds.Y + bounds.Height / 2;
        double sx = shape.LayoutTransform.M11, sy = shape.LayoutTransform.M22;
        // Layout is already expressed in physical units. Conjugate the child's
        // rotation into that basis, then apply the remaining ancestor map.
        return OfficeTransform.Translate(-x, -y).Then(OfficeTransform.Scale(1 / sx, 1 / sy))
            .Then(OfficeTransform.Scale(shape.Transform.FlipHorizontal ? -1 : 1, shape.Transform.FlipVertical ? -1 : 1))
            .Then(OfficeTransform.RotateDegrees(shape.Transform.RotationDegrees.GetValueOrDefault()))
            .Then(OfficeTransform.Scale(sx, sy)).Then(OfficeTransform.Translate(x, y)).Then(GroupPageTransform(shape));
    }

    private OfficeTransform GroupPageTransform(PublisherEscherShape shape) {
        if (shape.GroupTransform == shape.LayoutTransform) return OfficeTransform.Identity;
        var origin = PageOrigin(shape);
        return OfficeTransform.Translate(-origin.X, -origin.Y)
            .Then(OfficeTransform.Scale(12700, 12700))
            .Then(shape.LayoutTransform.Invert()).Then(shape.GroupTransform)
            .Then(OfficeTransform.Scale(1 / 12700D, 1 / 12700D))
            .Then(OfficeTransform.Translate(origin.X, origin.Y));
    }

    private readonly record struct FrameRectangle(double X, double Y, double Width, double Height);
}
