using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    /// <summary>Projects authored primitive geometry directly. A positive render canvas does not change a line's logical zero extent.</summary>
    private void ProjectAuthoredShape(PublisherShape shape, OfficeDrawing page) {
        OfficeShape geometry = shape.CreateShape();
        var local = new OfficeDrawing(Math.Max(shape.Width, 0.01), Math.Max(shape.Height, 0.01));
        if (HasFill(geometry) || geometry.StrokeColor.HasValue) local.AddShape(geometry, 0, 0);
        TagSourceElements(local, new[] { "publisher-object-" + shape.Id });
        if (shape.X < 0 || shape.Y < 0 || shape.X + shape.Width > page.Width || shape.Y + shape.Height > page.Height)
            _context.Add("PUB_PAGE_BLEED_CLIPPED", "Artwork outside the publication page remains in the scene but is clipped to the printable page in exports.",
                OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(shape.Id));
        _context.AccountProjection(local);
        if (local.Elements.Count != 0) page.AddDrawingForClippedRendering(local, shape.X, shape.Y, null);
        _projected.Add(shape.Id);
    }
}
