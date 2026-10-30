using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    /// <summary>Uses the semantic page registration recovered at the import boundary.</summary>
    private (double X, double Y) PageOrigin(PublisherEscherShape shape) {
        uint? owner = Owner(shape.Id, 0);
        return owner.HasValue && _pages.TryGetValue(owner.Value, out PublisherPageState? page)
            ? PageOrigin(page) : (_source.Width / 2, _source.Height / 2);
    }

    private static (double X, double Y) PageOrigin(PublisherPageState page) => (page.RegistrationX, page.RegistrationY);

    private FrameRectangle FrameBounds(PublisherEscherShape shape) {
        if (_publication.Pictures.TryGetValue(shape.Id, out PublisherPicture? picture) && picture.CanEdit)
            return new FrameRectangle(picture.X, picture.Y, picture.Width, picture.Height);
        var origin = PageOrigin(shape);
        var bounds = PublisherNativeGeometry.Bounds(shape, origin.X, origin.Y);
        return new FrameRectangle(bounds.X, bounds.Y, bounds.Width, bounds.Height);
    }

    private OfficeTransform PageTransform(PublisherEscherShape shape, FrameRectangle bounds) {
        if (_publication.Pictures.TryGetValue(shape.Id, out PublisherPicture? picture) && picture.CanEdit)
            return OfficeTransform.Translate(-bounds.X, -bounds.Y).Then(picture.PageTransform);
        var origin = PageOrigin(shape);
        return PublisherNativeGeometry.PageTransform(shape, bounds.X, bounds.Y, bounds.Width, bounds.Height, origin.X, origin.Y);
    }

    private OfficeTransform GroupPageTransform(PublisherEscherShape shape) {
        var origin = PageOrigin(shape);
        return PublisherNativeGeometry.GroupPageTransform(shape, origin.X, origin.Y);
    }

    private readonly record struct FrameRectangle(double X, double Y, double Width, double Height);
}
