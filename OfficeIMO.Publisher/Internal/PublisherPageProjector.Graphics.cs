using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private void ProjectGraphic(PublisherEscherShape source, OfficeDrawing drawing, double geometryWidth, double geometryHeight) {
        OfficeShape shape = Geometry(source, geometryWidth, geometryHeight);
        OfficeArtShapeStyle style = source.Style;
        shape.FillColor = style.FillEnabled != false ? Resolve(style.FillColor, source.Id) : null;
        shape.StrokeColor = style.LineEnabled != false ? Resolve(style.LineColor, source.Id) : null;
        shape.StrokeWidth = Math.Max(0, (style.LineWidthEmus ?? 9525) / 12700D);
        shape.FillOpacity = style.FillOpacity; shape.StrokeOpacity = style.LineOpacity;
        ProjectFill(source, shape, geometryWidth, geometryHeight);
        ProjectLineDetails(source, shape);
        if (style.HasProjectableShadow) {
            shape.Shadow = new OfficeShadow(Resolve(style.ShadowColor, source.Id) ?? OfficeColor.Black, style.ShadowOpacity ?? 1,
                (style.ShadowOffsetXEmus ?? 0) / 12700D, (style.ShadowOffsetYEmus ?? 0) / 12700D, Math.Max(0, (style.ShadowSoftnessEmus ?? 0) / 12700D));
        }
        uint? imageId = source.Property(0x104);
        if (!imageId.HasValue) {
            if (HasFill(shape) || shape.StrokeColor.HasValue || shape.Shadow != null) drawing.AddShape(shape, 0, 0);
            return;
        }
        // A picture paints inside its frame, before the outline. Keep fill and
        // shadow beneath it while preserving the native outline above it.
        OfficeShape background = shape.Clone();
        background.StrokeColor = null;
        if (HasFill(background) || background.Shadow != null) drawing.AddShape(background, 0, 0);
        ProjectPicture(source, drawing, imageId.Value);
        OfficeShape outline = shape.Clone();
        outline.FillColor = null; outline.FillGradient = null; outline.FillRadialGradient = null; outline.Shadow = null;
        if (outline.StrokeColor.HasValue) drawing.AddShape(outline, 0, 0);
    }
    private static bool HasFill(OfficeShape shape) => shape.FillColor.HasValue || shape.FillGradient != null || shape.FillRadialGradient != null;
    private void ProjectPicture(PublisherEscherShape source, OfficeDrawing drawing, uint imageId) {
        if (imageId > int.MaxValue || !_escher.Images.TryGetValue((int)imageId, out PublisherImage? image)) {
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
        drawing.AddImage(projectedBytes, projectedContentType, new OfficeImageProjection(new OfficeImagePlacement(0, 0, drawing.Width, drawing.Height), crop));
    }
    private OfficeShape Geometry(PublisherEscherShape source, double width, double height) {
        switch (source.Type) {
            case 1: case 75: case 202: return OfficeShape.Rectangle(width, height);
            case 2: return OfficeShape.RoundedRectangle(width, height, Math.Min(width, height) / 6);
            case 3: return OfficeShape.Ellipse(width, height);
            case 4: return OfficeShape.Polygon(new OfficePoint(width / 2, 0), new OfficePoint(width, height / 2), new OfficePoint(width / 2, height), new OfficePoint(0, height / 2));
            case 5: return OfficeShape.Polygon(new OfficePoint(width / 2, 0), new OfficePoint(width, height), new OfficePoint(0, height));
            case 6: return OfficeShape.Polygon(new OfficePoint(0, 0), new OfficePoint(width, height), new OfficePoint(0, height));
            case 20: return OfficeShape.Line(0, 0, width, height);
            default:
                _context.Add("PUB_SHAPE_GEOMETRY_APPROXIMATED", $"Native shape type {source.Type} was approximated by its bounding rectangle.", OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(source.Id));
                return OfficeShape.Rectangle(width, height);
        }
    }
    private OfficeColor? Resolve(OfficeArtColorReference? reference, uint id) {
        if (!reference.HasValue || reference.Value.IsIgnored) return null;
        if (reference.Value.TryResolve(index => index < _source.Palette.Count ? _source.Palette[index] : null, out OfficeColor color)) return color;
        _context.Add("PUB_SHAPE_COLOR_UNRESOLVED", "A native color reference was approximated as black.", OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(id));
        return OfficeColor.Black;
    }
}
