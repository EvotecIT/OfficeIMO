namespace OfficeIMO.Publisher.Internal;

/// <summary>Translates modern native page roles, ownership and picture bindings into the current publication model.</summary>
internal static partial class PublisherModelImporter {
    internal static PublisherPublicationState Import(PublisherContents source, PublisherQuillText text, PublisherEscherData artwork, PublisherParseContext context) {
        var owners = new Dictionary<uint, uint?>();
        var nativePages = source.Pages.ToDictionary(page => page.Id);
        uint? ResolveOwner(uint id, int depth) {
            context.CheckDepth(depth);
            if (owners.TryGetValue(id, out uint? owner)) return owner;
            if (source.ObjectPages.TryGetValue(id, out uint page)) owner = page;
            else if (source.Shapes.TryGetValue(id, out PublisherSourceShape? shape) && shape.Chunk.Parent.HasValue) {
                uint parent = shape.Chunk.Parent.Value;
                owner = nativePages.ContainsKey(parent) ? parent : ResolveOwner(parent, depth + 1);
            }
            owners[id] = owner;
            return owner;
        }
        foreach (uint id in source.Shapes.Keys.Concat(artwork.Shapes.Select(shape => shape.Id)).Distinct()) ResolveOwner(id, 0);
        var ordered = new Dictionary<uint, List<uint>>();
        foreach (uint id in artwork.Shapes.Select(shape => shape.Id).Concat(source.Shapes.Keys).Distinct()) {
            context.Token.ThrowIfCancellationRequested();
            uint? owner = owners[id];
            if (!owner.HasValue) continue;
            if (!ordered.TryGetValue(owner.Value, out List<uint>? ids)) ordered.Add(owner.Value, ids = new List<uint>());
            ids.Add(id);
        }
        PublisherPageState[] pages = source.Pages.Select(page => {
            double registrationX = source.Width / 2;
            bool qualified = !page.IsMaster || !page.NativeLayout.HasValue || page.NativeLayout.Value is 0 or 1 or 3;
            if (page.IsMaster) registrationX = page.NativeLayout switch { 0 => 0, 1 => source.Width, _ => registrationX };
            return new PublisherPageState(page.Id, page.Name, page.IsMaster, source.Width, source.Height,
                page.Master, registrationX, source.Height / 2, qualified, page.Background.GetValueOrDefault() != 0,
                ordered.TryGetValue(page.Id, out List<uint>? ids) ? ids : Array.Empty<uint>(),
                new PublisherFrameEditRange(registrationX + int.MinValue / 12700D, source.Height / 2 + int.MinValue / 12700D,
                    registrationX + int.MaxValue / 12700D, source.Height / 2 + int.MaxValue / 12700D));
        }).ToArray();
        var pageMap = pages.ToDictionary(page => page.Id);
        var pictures = new Dictionary<uint, PublisherPicture>();
        foreach (PublisherEscherShape shape in artwork.Shapes) {
            context.Token.ThrowIfCancellationRequested();
            if (!source.Shapes.TryGetValue(shape.Id, out PublisherSourceShape? native)) continue;
            if (!owners.TryGetValue(shape.Id, out uint? owner) || !owner.HasValue || !pageMap.TryGetValue(owner.Value, out PublisherPageState? page)) continue;
            PublisherPicture? picture = ImportPicture(shape, native, page, source, artwork);
            if (picture != null) pictures.Add(picture.Id, picture);
        }
        // Unavailable store entries and unresolved references still reserve their
        // identities. A new asset must never resolve an unrelated missing picture.
        int maximumImageId = Math.Max(context.ImageStoreEntryCount, artwork.Images.Keys.DefaultIfEmpty(0).Max());
        foreach (PublisherEscherShape shape in artwork.Shapes) {
            var reference = shape.Properties.LastOrDefault(property => property.PropertyId == 0x0104);
            if (reference is { IsComplex: false, IsBlipId: true, Value: > 0 and <= int.MaxValue })
                maximumImageId = Math.Max(maximumImageId, (int)reference.Value);
        }
        PublisherTextState currentText = ImportText(source, text, artwork, pageMap, owners, pictures, context);
        return new PublisherPublicationState(pages, owners, pictures, artwork.Images, currentText,
            nextImageId: maximumImageId == int.MaxValue ? 0 : maximumImageId + 1,
            nextObjectId: NextModelObjectId(source, text, owners, pages, currentText));
    }
}
