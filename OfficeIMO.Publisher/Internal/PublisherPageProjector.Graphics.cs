using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private void ProjectGraphic(PublisherEscherShape source, OfficeDrawing drawing, double geometryWidth, double geometryHeight) {
        if (_publication.Shapes.TryGetValue(source.Id, out PublisherShape? vector) && vector.CanEdit) {
            OfficeShape current = _publication.CurrentShape(source.Id).CreateShape();
            if (HasFill(current) || current.StrokeColor.HasValue) drawing.AddShape(current, 0, 0);
            return;
        }
        OfficeShape shape = new PublisherNativeShapeMapper(_source.Palette, _context).Map(source,
            geometryWidth, geometryHeight, drawing.Width, drawing.Height, out OfficeShape? clipShape);
        OfficeArtProperty? pictureReference = source.Properties.LastOrDefault(property => property.PropertyId == 0x0104);
        if (pictureReference == null || pictureReference.Value == 0 || pictureReference.IsComplex || !pictureReference.IsBlipId) {
            if (pictureReference?.IsComplex == true && pictureReference.Value != 0)
                _context.Add("PUB_IMAGE_DATA_UNSUPPORTED", "A native picture carries inline image data that is retained without artwork projection.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(source.Id));
            if (pictureReference is { IsComplex: false, IsBlipId: false, Value: not 0 })
                _context.Add("PUB_IMAGE_REFERENCE_INVALID", "A native picture reference lacks the required BLIP identifier flag; its image is retained without artwork projection.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(source.Id));
            if (HasFill(shape) || shape.StrokeColor.HasValue || shape.Shadow != null) drawing.AddShape(shape, 0, 0);
            return;
        }
        // A picture paints inside its frame, before the outline. Keep fill and
        // shadow beneath it while preserving the native outline above it.
        OfficeShape background = shape.Clone();
        background.StrokeColor = null;
        if (HasFill(background) || background.Shadow != null) drawing.AddShape(background, 0, 0);
        ProjectPicture(source, drawing, pictureReference.Value, clipShape);
        OfficeShape outline = shape.Clone();
        outline.FillColor = null; outline.FillGradient = null; outline.FillRadialGradient = null; outline.Shadow = null;
        if (outline.StrokeColor.HasValue) drawing.AddShape(outline, 0, 0);
    }
    private static bool HasFill(OfficeShape shape) => shape.FillColor.HasValue || shape.FillGradient != null || shape.FillRadialGradient != null;
    private void ProjectPicture(PublisherEscherShape source, OfficeDrawing drawing, uint imageId, OfficeShape? clipShape) {
        if (imageId > int.MaxValue || !_publication.Images.TryGetValue((int)imageId, out PublisherImage? image)) {
            _context.Add("PUB_IMAGE_REFERENCE_UNRESOLVED", "A native picture refers to an unavailable embedded image.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(source.Id)); return;
        }
        OfficeArtPictureProperties picture = OfficeArtPictureProperties.Decode(source.Properties);
        var crop = OfficeImageSourceCrop.FromClampedFractions(picture.CropFromLeft ?? 0, picture.CropFromTop ?? 0,
            picture.CropFromRight ?? 0, picture.CropFromBottom ?? 0);
        if (!crop.HasVisibleSourceArea) {
            _context.Add("PUB_IMAGE_CROP_INVALID", "The native crop removes the entire image; its picture is omitted.", OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(source.Id)); return;
        }
        OfficeArtPictureEffectProjector effects = PictureEffects(picture, source.Id);
        byte[] projectedBytes = ProjectImage(image, source.Id, effects, out string projectedContentType);
        var projection = new OfficeImageProjection(new OfficeImagePlacement(0, 0, drawing.Width, drawing.Height), crop);
        if (clipShape == null) drawing.AddImage(projectedBytes, projectedContentType, projection);
        else drawing.AddClippedImage(projectedBytes, projectedContentType, projection, 0, 0,
            OfficeClipPath.Path(drawing.Width, drawing.Height, clipShape.PathCommands, clipShape.FillRule));
    }
}
