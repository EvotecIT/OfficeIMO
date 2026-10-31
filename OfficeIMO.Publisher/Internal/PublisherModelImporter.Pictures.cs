using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherModelImporter {
    private static PublisherPicture? ImportPicture(PublisherEscherShape shape, PublisherSourceShape native,
        PublisherPageState page, PublisherContents source, PublisherEscherData artwork) {
        OfficeArtProperty? reference = shape.Properties.LastOrDefault(property => property.PropertyId == 0x0104);
        // Contents kind 0x20 also describes ordinary native lines. The
        // independently authored line control has this kind and no image
        // binding; it belongs to vector artwork rather than a missing picture.
        if (shape.Type == 20 && reference == null) return null;
        if (native.Chunk.Kind != 0x20 && shape.Type != 75 && reference == null) return null;
        int? imageId = reference is { IsComplex: false, IsBlipId: true, Value: > 0 and <= int.MaxValue } ? (int)reference.Value : null;
        OfficeArtPictureProperties picture = OfficeArtPictureProperties.Decode(shape.Properties);
        var crop = OfficeImageSourceCrop.FromClampedFractions(picture.CropFromLeft ?? 0, picture.CropFromTop ?? 0,
            picture.CropFromRight ?? 0, picture.CropFromBottom ?? 0);
        bool usableSpace = shape.LayoutTransform.TryInvert(out _) && shape.GroupTransform.TryInvert(out _);
        var bounds = shape.Bounds.HasValue ? PublisherNativeGeometry.Bounds(shape, page.RegistrationX, page.RegistrationY) : default;
        bool placement = shape.Bounds.HasValue && usableSpace && Finite(bounds.X) && Finite(bounds.Y)
            && Finite(bounds.Width) && Finite(bounds.Height) && bounds.Width > 0 && bounds.Height > 0;
        OfficeTransform transform = placement
            ? OfficeTransform.Translate(bounds.X, bounds.Y).Then(PublisherNativeGeometry.PageTransform(shape,
                bounds.X, bounds.Y, bounds.Width, bounds.Height, page.RegistrationX, page.RegistrationY))
            : OfficeTransform.Identity;
        string? limitation = PictureEditLimitation(shape, native, page, artwork, imageId, picture, crop, placement, source.Palette);
        OfficeShape? paint = limitation == null ? PicturePaint(shape, source.Palette) : null;
        if (paint != null) { paint.Width = bounds.Width; paint.Height = bounds.Height; }
        return new PublisherPicture(shape.Id, page.Id, imageId, bounds.X, bounds.Y, bounds.Width, bounds.Height,
            transform, crop, placement, limitation == null, paint, limitation);
    }

    private static string? PictureEditLimitation(PublisherEscherShape shape, PublisherSourceShape native, PublisherPageState page,
        PublisherEscherData artwork, int? imageId, OfficeArtPictureProperties picture, OfficeImageSourceCrop crop, bool placement,
        IReadOnlyList<OfficeColor?> palette) {
        if (!placement) return "The picture has no usable resolved page placement.";
        if (shape.Hidden) return "Hidden pictures are inspection-only.";
        if (shape.IsGroup || shape.HasPublicationGroupParent
            || shape.GroupTransform != OfficeTransform.Identity || shape.LayoutTransform != OfficeTransform.Identity)
            return "Grouped pictures and pictures using an enclosing drawing transform are inspection-only.";
        if (shape.Transform.RotationDegrees.GetValueOrDefault() != 0 || shape.Transform.FlipHorizontal || shape.Transform.FlipVertical)
            return "Rotated or reflected pictures are inspection-only.";
        if (!page.SupportsPlacementEditing) return "This master page's coordinate registration is not qualified for picture edits.";
        if (!imageId.HasValue || !artwork.Images.TryGetValue(imageId.Value, out PublisherImage? image))
            return "The picture has no resolved embedded image asset.";
        if (image.ContentType is not ("image/png" or "image/jpeg")) return "Picture edits currently require a PNG or JPEG asset.";
        if (native.Value(0x27).HasValue || native.Table != null) return "Pictures that also own text or table content are inspection-only.";
        if (shape.Type is not (1 or 75) || OfficeArtCustomPathProjector.HasPath(shape.Properties))
            return "Nonrectangular and custom-clipped pictures are inspection-only.";
        if (!crop.HasVisibleSourceArea || new[] { picture.CropFromLeft, picture.CropFromTop, picture.CropFromRight, picture.CropFromBottom }
            .Any(value => value.HasValue && (value.Value < 0 || value.Value >= 1)))
            return "The native crop needs fallback handling and is not qualified for picture edits.";
        OfficeArtPictureEffectProjector effects = OfficeArtPictureEffectProjector.Create(picture, _ => null);
        if (picture.HasPictureEffect || effects.HasProjection || effects.Limits != OfficeArtPictureEffectLimit.None)
            return "Pictures with native tone, transparency, recoloring or other image effects are inspection-only.";
        OfficeArtShapeStyle style = shape.Style;
        if (style.ShadowEnabled == true || style.HasUnprojectedVisualStyle
            || style.LineEnabled != false && (style.LineDashing.GetValueOrDefault() != 0
                || style.LineStartArrowhead.GetValueOrDefault() != 0 || style.LineEndArrowhead.GetValueOrDefault() != 0
                || style.LineEndCapStyle > 2 || style.LineJoinStyle > 2
                || style.LineJoinStyle == 1 && style.LineMiterLimit < 1))
            return "The picture's fill, outline or shadow requires an unsupported editable paint model.";
        foreach (OfficeArtColorReference? color in new[] { style.FillEnabled == false ? null : style.FillColor, style.LineEnabled == false ? null : style.LineColor })
            if (color.HasValue && !color.Value.IsIgnored && !color.Value.TryResolve(index => index < palette.Count ? palette[index] : null, out _))
                return "The picture's paint contains an unresolved color reference.";
        return null;
    }

    private static OfficeShape PicturePaint(PublisherEscherShape shape, IReadOnlyList<OfficeColor?> palette) {
        OfficeArtShapeStyle style = shape.Style;
        OfficeColor? Resolve(OfficeArtColorReference? reference) {
            if (!reference.HasValue || reference.Value.IsIgnored) return null;
            return reference.Value.TryResolve(index => index < palette.Count ? palette[index] : null, out OfficeColor color) ? color : OfficeColor.Black;
        }
        var result = new OfficeShape {
            Kind = OfficeShapeKind.Rectangle,
            FillColor = style.FillEnabled != false && !OfficeArtCustomPathProjector.IsFillDisabled(shape.Properties) ? Resolve(style.FillColor) : null,
            StrokeColor = style.LineEnabled != false && !OfficeArtCustomPathProjector.IsLineDisabled(shape.Properties) ? Resolve(style.LineColor) : null,
            StrokeWidth = Math.Max(0, (style.LineWidthEmus ?? 9525) / 12700D),
            FillOpacity = style.FillOpacity, StrokeOpacity = style.LineOpacity
        };
        if (result.StrokeWidth > 0 && result.StrokeColor.HasValue) {
            result.StrokeLineCap = style.LineEndCapStyle switch { 0 => OfficeStrokeLineCap.Round, 1 => OfficeStrokeLineCap.Square, _ => OfficeStrokeLineCap.Butt };
            result.StrokeLineJoin = style.LineJoinStyle switch { 0 => OfficeStrokeLineJoin.Bevel, 1 => OfficeStrokeLineJoin.Miter, _ => OfficeStrokeLineJoin.Round };
            result.StrokeMiterLimit = style.LineMiterLimit >= 1 ? style.LineMiterLimit.Value : 8;
        }
        return result;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
