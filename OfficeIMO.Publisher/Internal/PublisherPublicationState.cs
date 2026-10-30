using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Generation-neutral current page ownership, placed pictures and embedded assets. Native bindings are retained by the importer separately.</summary>
internal sealed partial class PublisherPublicationState {
    private readonly HashSet<int> _sourceImageIds;
    internal PublisherPublicationState(IEnumerable<PublisherPageState> pages, IReadOnlyDictionary<uint, uint?> owners,
        IReadOnlyDictionary<uint, PublisherPicture> pictures, IReadOnlyDictionary<int, PublisherImage> images, PublisherTextState text,
        IReadOnlyDictionary<uint, uint?> tableStoryReferences,
        IEnumerable<int>? sourceImageIds = null, int? nextImageId = null, uint nextObjectId = 0, IReadOnlyDictionary<uint, PublisherShape>? shapes = null,
        IReadOnlyDictionary<uint, PublisherTableState>? tables = null) {
        Pages = Array.AsReadOnly(pages.ToArray());
        ObjectOwners = new System.Collections.ObjectModel.ReadOnlyDictionary<uint, uint?>(owners.ToDictionary(item => item.Key, item => item.Value));
        Pictures = new System.Collections.ObjectModel.ReadOnlyDictionary<uint, PublisherPicture>(pictures.ToDictionary(item => item.Key, item => item.Value));
        Images = new System.Collections.ObjectModel.ReadOnlyDictionary<int, PublisherImage>(images.ToDictionary(item => item.Key, item => item.Value));
        Text = text;
        Shapes = new System.Collections.ObjectModel.ReadOnlyDictionary<uint, PublisherShape>((shapes ?? new Dictionary<uint, PublisherShape>()).ToDictionary(item => item.Key, item => item.Value));
        Tables = new System.Collections.ObjectModel.ReadOnlyDictionary<uint, PublisherTableState>((tables ?? new Dictionary<uint, PublisherTableState>()).ToDictionary(item => item.Key, item => item.Value));
        TableStoryReferences = new System.Collections.ObjectModel.ReadOnlyDictionary<uint, uint?>(tableStoryReferences.ToDictionary(item => item.Key, item => item.Value));
        _sourceImageIds = new HashSet<int>(sourceImageIds ?? images.Keys);
        NextImageId = nextImageId ?? (images.Count == 0 ? 1 : images.Keys.Max() == int.MaxValue ? 0 : images.Keys.Max() + 1);
        NextObjectId = nextObjectId;
    }
    internal IReadOnlyList<PublisherPageState> Pages { get; }
    internal IReadOnlyDictionary<uint, uint?> ObjectOwners { get; }
    internal IReadOnlyDictionary<uint, PublisherPicture> Pictures { get; }
    internal IReadOnlyDictionary<int, PublisherImage> Images { get; }
    internal PublisherTextState Text { get; }
    internal IReadOnlyDictionary<uint, PublisherShape> Shapes { get; }
    internal IReadOnlyDictionary<uint, PublisherTableState> Tables { get; }
    /// <summary>Current table-to-story ownership, including imported occurrences whose grid or page cannot be resolved. Admission must account for these even when no placed table snapshot can be produced.</summary>
    internal IReadOnlyDictionary<uint, uint?> TableStoryReferences { get; }
    internal int AuthoredObjectCount => Text.AuthoredFrameCount + Shapes.Values.Count(shape => shape.IsAuthored) + Tables.Values.Count(table => table.IsAuthored);
    internal int NextImageId { get; }
    internal uint NextObjectId { get; }
    internal long AdditionalImageBytes => Images.Values.Where(image => !_sourceImageIds.Contains(image.Id)).Sum(image => (long)image.ByteCount);
    internal int AdditionalImageCount => Images.Keys.Count(id => !_sourceImageIds.Contains(id));

    internal PublisherPicture EditablePicture(uint pictureId) {
        if (!Pictures.TryGetValue(pictureId, out PublisherPicture? picture))
            throw new ArgumentOutOfRangeException(nameof(pictureId), "The publication has no such picture.");
        if (!picture.CanEdit) throw new NotSupportedException("Only visible, resolved, untransformed top-level PNG/JPEG pictures with qualified rectangular paint and crop can be edited.");
        return picture;
    }

    internal PublisherPublicationState SetBounds(PublisherPicture picture, double x, double y, double width, double height) {
        var pictures = Pictures.ToDictionary(item => item.Key, item => item.Value);
        pictures[picture.Id] = picture.WithBounds(x, y, width, height);
        return new PublisherPublicationState(Pages, ObjectOwners, pictures, Images, Text, TableStoryReferences, _sourceImageIds, NextImageId, NextObjectId, Shapes, Tables);
    }

    internal PublisherPublicationState ReplaceImage(PublisherPicture picture, byte[] bytes, string contentType) {
        if (NextImageId == 0) throw new InvalidOperationException("The publication image identifier space is exhausted.");
        var pictures = Pictures.ToDictionary(item => item.Key, item => item.Value);
        pictures[picture.Id] = picture.WithImage(NextImageId);
        var images = Images.ToDictionary(item => item.Key, item => item.Value);
        images.Add(NextImageId, new PublisherImage(NextImageId, contentType, bytes));
        // Keep every original native asset. Superseded edit-owned bytes have no
        // preservation role and can be released when no current placement uses them.
        if (picture.ImageId.HasValue && !_sourceImageIds.Contains(picture.ImageId.Value)
            && !pictures.Values.Any(item => item.ImageId == picture.ImageId)) images.Remove(picture.ImageId.Value);
        return new PublisherPublicationState(Pages, ObjectOwners, pictures, images, Text, TableStoryReferences, _sourceImageIds,
            NextImageId == int.MaxValue ? 0 : NextImageId + 1, NextObjectId, Shapes, Tables);
    }

    internal PublisherPublicationState WithText(PublisherTextState text) =>
        new(Pages, ObjectOwners, Pictures, Images, text, TableStoryReferences, _sourceImageIds, NextImageId, NextObjectId, Shapes, Tables);

    internal PublisherPublicationState AddLinkedTextFrame(uint previousFrameId, uint pageId,
        double x, double y, double width, double height, out uint frameId) {
        if (!Text.Frames.TryGetValue(previousFrameId, out PublisherFrameState? previous))
            throw new ArgumentOutOfRangeException(nameof(previousFrameId), "The publication has no such text frame.");
        PublisherPageState? page = Pages.SingleOrDefault(item => item.Id == pageId);
        if (page == null) throw new ArgumentOutOfRangeException(nameof(pageId), "The publication has no such page.");
        if (page.IsMaster) throw new NotSupportedException("A continuation frame must belong to a document page.");
        if (Text.IsTableStory(previous.StoryId) || !Text.Stories.ContainsKey(previous.StoryId)
            || !Text.Chains.TryGetValue(previous.StoryId, out IReadOnlyList<uint>? chain) || chain.Count == 0 || chain[chain.Count - 1] != previousFrameId)
            throw new NotSupportedException("A continuation frame requires the tail of a resolved non-table story chain.");
        var pages = Pages.ToDictionary(item => item.Id);
        if (chain.Any(id => !Text.Frames[id].Available || !Text.Frames[id].SupportsContinuation
            || !Text.Frames[id].PageId.HasValue || !pages.TryGetValue(Text.Frames[id].PageId!.Value, out PublisherPageState? owner) || owner.IsMaster))
            throw new NotSupportedException("Continuation creation requires a complete chain of ordinary, visible, untransformed document-page text frames.");
        if (NextObjectId == 0) throw new InvalidOperationException("The publication object identifier space is exhausted.");
        frameId = NextObjectId;
        PublisherTextState text = Text.AppendFrame(previous, page, x, y, width, height, frameId);
        var owners = ObjectOwners.ToDictionary(item => item.Key, item => item.Value); owners.Add(frameId, pageId);
        uint createdId = frameId;
        return new PublisherPublicationState(Pages.Select(item => item.Id == pageId ? item.WithObject(createdId) : item),
            owners, Pictures, Images, text, TableStoryReferences, _sourceImageIds, NextImageId, createdId == uint.MaxValue ? 0 : createdId + 1, Shapes, Tables);
    }
}

/// <summary>A page definition and its object order. Registration is the semantic page basis used when composing an applied master.</summary>
internal sealed class PublisherPageState {
    internal PublisherPageState(uint id, string name, bool isMaster, double width, double height, uint? masterPageId,
        double registrationX, double registrationY, bool supportsPlacementEditing, bool hasUnassessedBackground,
        IReadOnlyList<uint> objectIds, PublisherFrameEditRange frameEditRange) {
        Id = id; Name = name; IsMaster = isMaster; Width = width; Height = height; MasterPageId = masterPageId;
        RegistrationX = registrationX; RegistrationY = registrationY; SupportsPlacementEditing = supportsPlacementEditing;
        HasUnassessedBackground = hasUnassessedBackground; ObjectIds = Array.AsReadOnly(objectIds.ToArray());
        FrameEditRange = frameEditRange;
    }
    internal uint Id { get; }
    internal string Name { get; }
    internal bool IsMaster { get; }
    internal double Width { get; }
    internal double Height { get; }
    internal uint? MasterPageId { get; }
    internal double RegistrationX { get; }
    internal double RegistrationY { get; }
    internal bool SupportsPlacementEditing { get; }
    internal bool HasUnassessedBackground { get; }
    internal IReadOnlyList<uint> ObjectIds { get; }
    internal PublisherFrameEditRange FrameEditRange { get; }
    internal PublisherPageState WithObject(uint id) => new(Id, Name, IsMaster, Width, Height, MasterPageId,
        RegistrationX, RegistrationY, SupportsPlacementEditing, HasUnassessedBackground,
        ObjectIds.Concat(new[] { id }).ToArray(), FrameEditRange);
}
