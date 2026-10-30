using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    /// <summary>Projects the current placed-picture model directly, without constructing or editing native drawing records.</summary>
    private void ProjectSemanticPicture(PublisherPicture picture, OfficeDrawing page) {
        OfficeShape paint = picture.CreatePaint();
        var local = new OfficeDrawing(picture.Width, picture.Height);
        OfficeShape background = paint.Clone(); background.StrokeColor = null;
        if (HasFill(background)) local.AddShape(background, 0, 0);
        PublisherImage asset = _publication.Images[picture.ImageId!.Value];
        local.AddImage(asset.Bytes, asset.ContentType,
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, picture.Width, picture.Height), picture.Crop));
        OfficeShape outline = paint.Clone(); outline.FillColor = null;
        if (outline.StrokeColor.HasValue) local.AddShape(outline, 0, 0);
        TagSourceElements(local, new[] { "publisher-object-" + picture.Id });
        if (picture.X < 0 || picture.Y < 0 || picture.X + picture.Width > page.Width || picture.Y + picture.Height > page.Height)
            _context.Add("PUB_PAGE_BLEED_CLIPPED", "Artwork outside the publication page remains in the scene but is clipped to the printable page in exports.",
                OfficeConversionLossKind.Approximation, picture.IsAuthored ? "publication/picture/" + picture.Id : PublisherEscherReader.ShapeLocation(picture.Id));
        _context.AccountProjection(local);
        page.AddDrawingForClippedRendering(local, picture.X, picture.Y,
            new OfficeImageFrameTransform(0, picture.X + picture.Width / 2, picture.Y + picture.Height / 2, false, false));
        _projected.Add(picture.Id);
    }
}
