using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Generation-neutral current page ownership, placed pictures and embedded assets. Native bindings are retained by the importer separately.</summary>
internal sealed class PublisherPublicationState {
    private readonly HashSet<int> _sourceImageIds;
    internal PublisherPublicationState(IEnumerable<PublisherPageState> pages, IReadOnlyDictionary<uint, uint?> owners,
        IReadOnlyDictionary<uint, PublisherPicture> pictures, IReadOnlyDictionary<int, PublisherImage> images,
        IEnumerable<int>? sourceImageIds = null, int? nextImageId = null) {
        Pages = Array.AsReadOnly(pages.ToArray());
        ObjectOwners = new System.Collections.ObjectModel.ReadOnlyDictionary<uint, uint?>(owners.ToDictionary(item => item.Key, item => item.Value));
        Pictures = new System.Collections.ObjectModel.ReadOnlyDictionary<uint, PublisherPicture>(pictures.ToDictionary(item => item.Key, item => item.Value));
        Images = new System.Collections.ObjectModel.ReadOnlyDictionary<int, PublisherImage>(images.ToDictionary(item => item.Key, item => item.Value));
        _sourceImageIds = new HashSet<int>(sourceImageIds ?? images.Keys);
        NextImageId = nextImageId ?? (images.Count == 0 ? 1 : images.Keys.Max() == int.MaxValue ? 0 : images.Keys.Max() + 1);
    }
    internal IReadOnlyList<PublisherPageState> Pages { get; }
    internal IReadOnlyDictionary<uint, uint?> ObjectOwners { get; }
    internal IReadOnlyDictionary<uint, PublisherPicture> Pictures { get; }
    internal IReadOnlyDictionary<int, PublisherImage> Images { get; }
    internal int NextImageId { get; }
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
        return new PublisherPublicationState(Pages, ObjectOwners, pictures, Images, _sourceImageIds, NextImageId);
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
        return new PublisherPublicationState(Pages, ObjectOwners, pictures, images, _sourceImageIds,
            NextImageId == int.MaxValue ? 0 : NextImageId + 1);
    }
}

/// <summary>A page definition and its object order. Registration is the semantic page basis used when composing an applied master.</summary>
internal sealed class PublisherPageState {
    internal PublisherPageState(uint id, string name, bool isMaster, double width, double height, uint? masterPageId,
        double registrationX, double registrationY, bool supportsPlacementEditing, bool hasUnassessedBackground,
        IReadOnlyList<uint> objectIds) {
        Id = id; Name = name; IsMaster = isMaster; Width = width; Height = height; MasterPageId = masterPageId;
        RegistrationX = registrationX; RegistrationY = registrationY; SupportsPlacementEditing = supportsPlacementEditing;
        HasUnassessedBackground = hasUnassessedBackground; ObjectIds = Array.AsReadOnly(objectIds.ToArray());
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
}
