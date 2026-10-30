using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPublicationState {
    internal PublisherPageState PicturePage(uint pageId, double x, double y, double width, double height) {
        PublisherPageState? page = Pages.SingleOrDefault(item => item.Id == pageId);
        if (page == null) throw new ArgumentOutOfRangeException(nameof(pageId), "The publication has no such page.");
        if (!page.SupportsPlacementEditing) throw new NotSupportedException("This page's coordinate registration is not qualified for picture creation.");
        page.FrameEditRange.Validate(x, y, width, height);
        return page;
    }

    internal PublisherPublicationState AddPicture(PublisherPageState page, byte[] bytes, string contentType,
        double x, double y, double width, double height, out uint pictureId) {
        if (NextObjectId == 0) throw new InvalidOperationException("The publication object identifier space is exhausted.");
        if (NextImageId == 0) throw new InvalidOperationException("The publication image identifier space is exhausted.");
        pictureId = NextObjectId;
        var pictures = Pictures.ToDictionary(item => item.Key, item => item.Value);
        var paint = new OfficeShape { Kind = OfficeShapeKind.Rectangle, Width = width, Height = height };
        pictures.Add(pictureId, new PublisherPicture(pictureId, page.Id, NextImageId, x, y, width, height,
            OfficeTransform.Translate(x, y), default, true, true, paint, isAuthored: true));
        var images = Images.ToDictionary(item => item.Key, item => item.Value);
        images.Add(NextImageId, new PublisherImage(NextImageId, contentType, bytes));
        var owners = ObjectOwners.ToDictionary(item => item.Key, item => item.Value); owners.Add(pictureId, page.Id);
        uint createdId = pictureId;
        return new PublisherPublicationState(Pages.Select(item => item.Id == page.Id ? item.WithObject(createdId) : item),
            owners, pictures, images, Text, TableStoryReferences, _sourceImageIds,
            NextImageId == int.MaxValue ? 0 : NextImageId + 1, createdId == uint.MaxValue ? 0 : createdId + 1, Shapes, Tables);
    }
}
