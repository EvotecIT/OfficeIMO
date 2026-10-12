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
        shape.StrokeDashStyle = style.LineDashing switch { 1 or 5 => OfficeStrokeDashStyle.Dash, 2 or 6 => OfficeStrokeDashStyle.Dot,
            3 or 7 => OfficeStrokeDashStyle.DashDot, 4 or 8 => OfficeStrokeDashStyle.DashDotDot, _ => OfficeStrokeDashStyle.Solid };
        if (style.HasProjectableShadow) {
            shape.Shadow = new OfficeShadow(Resolve(style.ShadowColor, source.Id) ?? OfficeColor.Black, style.ShadowOpacity ?? 1,
                (style.ShadowOffsetXEmus ?? 0) / 12700D, (style.ShadowOffsetYEmus ?? 0) / 12700D, Math.Max(0, (style.ShadowSoftnessEmus ?? 0) / 12700D));
        }
        if (style.FillEnabled != false && style.FillType.HasValue && style.FillType > 0) {
            _context.Add("PUB_FILL_APPROXIMATED", "A non-solid native fill was approximated by its primary color.", OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(source.Id));
        }
        if ((style.LineStartArrowhead ?? 0) != 0 || (style.LineEndArrowhead ?? 0) != 0 || (style.LineStyle ?? 0) != 0)
            _context.Add("PUB_LINE_DETAIL_UNASSESSED", "Native arrowheads and compound line styles require additional projection.", OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(source.Id));
        uint? imageId = source.Property(0x104);
        if (!imageId.HasValue) {
            if (shape.FillColor.HasValue || shape.StrokeColor.HasValue || shape.Shadow != null) drawing.AddShape(shape, 0, 0);
            return;
        }
        // A picture paints inside its frame, before the outline. Keep fill and
        // shadow beneath it while preserving the native outline above it.
        OfficeShape background = shape.Clone();
        background.StrokeColor = null;
        if (background.FillColor.HasValue || background.Shadow != null) drawing.AddShape(background, 0, 0);
        ProjectPicture(source, drawing, imageId.Value);
        OfficeShape outline = shape.Clone();
        outline.FillColor = null; outline.Shadow = null;
        if (outline.StrokeColor.HasValue) drawing.AddShape(outline, 0, 0);
    }
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
        byte[] projectedBytes = ProjectImage(image, source.Id, out string projectedContentType);
        drawing.AddImage(projectedBytes, projectedContentType, new OfficeImageProjection(new OfficeImagePlacement(0, 0, drawing.Width, drawing.Height), crop));
        if (picture.HasPictureEffect) _context.Add("PUB_PICTURE_EFFECT_UNASSESSED", "Native picture recoloring, brightness, contrast, or transparency effects require additional projection.",
            OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(source.Id));
    }
    private byte[] ProjectImage(PublisherImage image, uint objectId, out string contentType) {
        contentType = image.ContentType;
        OfficeImageFormat format = OfficeImageInfo.FromMimeType(contentType);
        if (format is not (OfficeImageFormat.Wmf or OfficeImageFormat.Emf)) return image.Bytes;
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var codec = new OfficeRasterImageFallbackCodec(_context.Options.ImageCodec, diagnostics, PublisherEscherReader.ShapeLocation(objectId));
        _context.Token.ThrowIfCancellationRequested();
        codec.TryDecode(image.Bytes, contentType, out OfficeRasterImage? raster);
        _context.Token.ThrowIfCancellationRequested();
        if (raster == null || (long)raster.Width * raster.Height > _context.Options.MaximumRasterPixels)
            throw new InvalidDataException("Publisher application image codec exceeded the pixel limit or returned no image.");
        foreach (OfficeImageExportDiagnostic diagnostic in diagnostics) _context.Add(diagnostic.Code, diagnostic.Message, diagnostic.LossKind, diagnostic.Source);
        if (!diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Omission))
            _context.Add("PUB_METAFILE_RASTERIZED", "A WMF/EMF picture was rasterized by the application codec. Images retains the original native payload.",
                OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(objectId));
        contentType = "image/png";
        return OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png, null, _context.Options.MaximumImageBytes, _context.Token);
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
