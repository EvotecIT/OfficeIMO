using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherModelImporter {
    private static PublisherTextState ImportText(PublisherContents source, PublisherQuillText text, PublisherEscherData artwork,
        IReadOnlyDictionary<uint, PublisherPageState> pages, IReadOnlyDictionary<uint, uint?> owners,
        IReadOnlyDictionary<uint, PublisherPicture> pictures, PublisherParseContext context) {
        var shapes = artwork.Shapes.ToDictionary(shape => shape.Id);
        var frames = new Dictionary<uint, PublisherFrameState>();
        foreach (PublisherSourceShape native in source.Shapes.Values.Where(shape => shape.Chunk.Kind != 0x10 && shape.Value(0x27).HasValue)) {
            context.Token.ThrowIfCancellationRequested();
            owners.TryGetValue(native.Chunk.Id, out uint? owner);
            PublisherPageState? page = owner.HasValue && pages.TryGetValue(owner.Value, out PublisherPageState? found) ? found : null;
            shapes.TryGetValue(native.Chunk.Id, out PublisherEscherShape? shape);
            double originX = page?.RegistrationX ?? source.Width / 2, originY = page?.RegistrationY ?? source.Height / 2;
            bool anchor = shape?.Bounds.HasValue == true;
            bool transformed = shape != null && shape.LayoutTransform.TryInvert(out _) && shape.GroupTransform.TryInvert(out _);
            var bounds = anchor ? PublisherNativeGeometry.Bounds(shape!, originX, originY) : default;
            OfficeTransform transform = anchor && transformed
                ? OfficeTransform.Translate(bounds.X, bounds.Y).Then(PublisherNativeGeometry.PageTransform(shape!,
                    bounds.X, bounds.Y, bounds.Width, bounds.Height, originX, originY)) : OfficeTransform.Identity;
            bool editable = page?.SupportsPlacementEditing == true && anchor && shape != null && !shape.Hidden && !shape.IsGroup
                && !shape.HasPublicationGroupParent && shape.GroupTransform == OfficeTransform.Identity && shape.LayoutTransform == OfficeTransform.Identity
                && shape.Transform.RotationDegrees.GetValueOrDefault() == 0 && !shape.Transform.FlipHorizontal && !shape.Transform.FlipVertical;
            uint columns = shape?.Property(0x8C) ?? 1;
            double gap = (shape?.Property(0x8D) ?? 0) / 12700D;
            var padding = new OfficeTextPadding((shape?.Property(0x81) ?? 36576) / 12700D, (shape?.Property(0x82) ?? 36576) / 12700D,
                (shape?.Property(0x83) ?? 36576) / 12700D, (shape?.Property(0x84) ?? 36576) / 12700D);
            uint vertical = native.Value(0x35) ?? 0;
            bool ordinary = editable && (shape!.Type is 1 or 202) && !OfficeArtCustomPathProjector.HasPath(shape.Properties)
                && !pictures.ContainsKey(shape.Id);
            frames.Add(native.Chunk.Id, new PublisherFrameState(native.Chunk.Id, native.Value(0x27)!.Value, owner,
                bounds.X, bounds.Y, bounds.Width, bounds.Height, transform, anchor, transformed, shape?.Hidden == true,
                shape != null && !shape.IsGroup, editable, ordinary, columns > int.MaxValue ? int.MaxValue : (int)columns, gap, padding,
                vertical switch { 1 => OfficeTextVerticalAlignment.Center, 2 => OfficeTextVerticalAlignment.Bottom, _ => OfficeTextVerticalAlignment.Top },
                native.WrapObjects, page?.FrameEditRange ?? default));
        }
        IReadOnlyDictionary<uint, IReadOnlyList<uint>> chains = ImportFrameChains(source, text, frames, context);
        var margins = new Dictionary<uint, PublisherWrapInfo>();
        var obstacles = new Dictionary<uint, PublisherImportedWrapObject>();
        foreach (PublisherEscherShape shape in artwork.Shapes) {
            context.Token.ThrowIfCancellationRequested();
            double[] distances = new[] { 0x384, 0x385, 0x386, 0x387 }.Select(property => (double)unchecked((int)(shape.Property(property) ?? 0))).ToArray();
            margins.Add(shape.Id, new PublisherWrapInfo(new OfficeTextPadding(Math.Max(0, distances[0]) / 12700D, Math.Max(0, distances[1]) / 12700D,
                Math.Max(0, distances[2]) / 12700D, Math.Max(0, distances[3]) / 12700D), distances.Any(value => value < 0),
                shape.Properties.Any(property => property.PropertyId == 0x383)));
            if (frames.ContainsKey(shape.Id) || pictures.TryGetValue(shape.Id, out PublisherPicture? picture) && picture.CanEdit) continue;
            owners.TryGetValue(shape.Id, out uint? owner);
            PublisherPageState? page = owner.HasValue && pages.TryGetValue(owner.Value, out PublisherPageState? found) ? found : null;
            double originX = page?.RegistrationX ?? source.Width / 2, originY = page?.RegistrationY ?? source.Height / 2;
            bool available = shape.Bounds.HasValue && shape.LayoutTransform.TryInvert(out _) && shape.GroupTransform.TryInvert(out _);
            var bounds = shape.Bounds.HasValue ? PublisherNativeGeometry.Bounds(shape, originX, originY) : default;
            var transform = available ? OfficeTransform.Translate(bounds.X, bounds.Y).Then(PublisherNativeGeometry.PageTransform(shape,
                bounds.X, bounds.Y, bounds.Width, bounds.Height, originX, originY)) : OfficeTransform.Identity;
            obstacles.Add(shape.Id, new PublisherImportedWrapObject(owner, bounds.X, bounds.Y, bounds.Width, bounds.Height, transform, available, shape.Hidden));
        }
        return new PublisherTextState(text.Stories, frames, chains,
            text.CellParagraphs.Keys.Concat(source.Shapes.Values.Where(shape => shape.Chunk.Kind == 0x10 && shape.Value(0x27).HasValue).Select(shape => shape.Value(0x27)!.Value)),
            text.CellParagraphs, margins, obstacles);
    }
}
