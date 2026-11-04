using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherModelImporter {
    private static IReadOnlyDictionary<uint, PublisherShape> ImportShapes(PublisherContents source, PublisherEscherData artwork,
        IReadOnlyDictionary<uint, PublisherPageState> pages, IReadOnlyDictionary<uint, uint?> owners,
        IReadOnlyDictionary<uint, PublisherPicture> pictures, PublisherTextState text, PublisherParseContext context) {
        var result = new Dictionary<uint, PublisherShape>();
        var mapper = new PublisherNativeShapeMapper(source.Palette, context);
        foreach (PublisherEscherShape shape in artwork.Shapes) {
            context.Token.ThrowIfCancellationRequested();
            if (shape.IsGroup || pictures.ContainsKey(shape.Id)
                || !source.Shapes.TryGetValue(shape.Id, out PublisherSourceShape? native) || native.Table != null || native.Chunk.Kind == 0x10
                || !owners.TryGetValue(shape.Id, out uint? owner) || !owner.HasValue
                || !pages.TryGetValue(owner.Value, out PublisherPageState? page)) continue;
            bool space = shape.LayoutTransform.TryInvert(out _) && shape.GroupTransform.TryInvert(out _);
            var bounds = shape.Bounds.HasValue ? PublisherNativeGeometry.Bounds(shape, page.RegistrationX, page.RegistrationY) : default;
            // Custom paths can represent a legitimate horizontal or vertical stroke.
            // Recover that placement independently of primitive edit admission.
            bool oneDimensional = shape.Type == 20 || OfficeArtCustomPathProjector.HasPath(shape.Properties);
            bool placement = shape.Bounds.HasValue && space && Finite(bounds.X) && Finite(bounds.Y) && Finite(bounds.Width) && Finite(bounds.Height)
                && bounds.Width >= 0 && bounds.Height >= 0 && (oneDimensional ? bounds.Width > 0 || bounds.Height > 0 : bounds.Width > 0 && bounds.Height > 0);
            OfficeTransform transform = placement ? OfficeTransform.Translate(bounds.X, bounds.Y).Then(PublisherNativeGeometry.PageTransform(shape,
                bounds.X, bounds.Y, bounds.Width, bounds.Height, page.RegistrationX, page.RegistrationY)) : OfficeTransform.Identity;
            string? limitation = ShapeEditLimitation(shape, page, source.Palette, placement);
            if (limitation == null && text.Frames.TryGetValue(shape.Id, out PublisherFrameState? frame) && !frame.CanEditBounds)
                limitation = "The shape's text-frame placement is not editable.";
            OfficeShape? geometry = limitation == null ? mapper.Map(shape, bounds.Width, bounds.Height, bounds.Width, bounds.Height, out _) : null;
            result.Add(shape.Id, new PublisherShape(shape.Id, page.Id, bounds.X, bounds.Y, bounds.Width, bounds.Height,
                transform, placement, geometry, limitation));
        }
        return result;
    }

    private static string? ShapeEditLimitation(PublisherEscherShape shape, PublisherPageState page,
        IReadOnlyList<OfficeColor?> palette, bool placement, bool allowTableFrame = false) {
        if (!placement) return "The artwork has no usable resolved page placement.";
        if (shape.Hidden) return "Hidden artwork is inspection-only.";
        if (shape.HasPublicationGroupParent || shape.GroupTransform != OfficeTransform.Identity || shape.LayoutTransform != OfficeTransform.Identity
            || shape.Transform.RotationDegrees.GetValueOrDefault() != 0 || shape.Transform.FlipHorizontal || shape.Transform.FlipVertical)
            return "Grouped, rotated and reflected artwork is inspection-only.";
        if (!page.SupportsPlacementEditing) return "This master page's coordinate registration is not qualified for shape edits.";
        if (shape.Type is not (1 or 2 or 3 or 4 or 5 or 6 or 20 or 202) && !(allowTableFrame && (shape.Type is 75 or 201))
            || OfficeArtCustomPathProjector.HasPath(shape.Properties))
            return "The artwork's preset or custom geometry has no qualified editable model.";
        OfficeArtShapeStyle style = shape.Style;
        // This is an in-memory solid-paint contract, not a native rewrite gate.
        // Ordinary Publisher shapes carry unused fill/line background colors;
        // those do not change their solid rendered paint.
        if (style.ShadowEnabled == true || style.HasUnprojectedVisualStyle
            || shape.Properties.Any(property => property.PropertyId is >= 0x0280 and <= 0x02FF && property.Value != 0)
            || style.LineEnabled != false && (style.LineDashing.GetValueOrDefault() != 0 || style.LineStartArrowhead.GetValueOrDefault() != 0
                || style.LineEndArrowhead.GetValueOrDefault() != 0 || style.LineEndCapStyle > 2 || style.LineJoinStyle > 2
                || style.LineJoinStyle == 1 && style.LineMiterLimit < 1))
            return "The artwork's gradient, effects, decorations or other paint require an unsupported editable model.";
        foreach (OfficeArtColorReference? color in new[] { style.FillEnabled == false ? null : style.FillColor, style.LineEnabled == false ? null : style.LineColor })
            if (color.HasValue && !color.Value.IsIgnored && !color.Value.TryResolve(index => index < palette.Count ? palette[index] : null, out _))
                return "The artwork has an unresolved paint color.";
        if (style.LineWidthEmus < 0 || style.FillOpacity is < 0 or > 1 || style.LineOpacity is < 0 or > 1)
            return "The artwork's native paint values require fallback handling.";
        return null;
    }
}
