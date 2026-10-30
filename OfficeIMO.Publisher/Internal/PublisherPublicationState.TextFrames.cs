using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPublicationState {
    internal PublisherPublicationState AddTextFrame(uint pageId, double x, double y, double width, double height,
        IReadOnlyList<OfficeRichTextParagraph> paragraphs, PublisherReadOptions options, CancellationToken token, out uint frameId) {
        PublisherPageState? page = Pages.SingleOrDefault(item => item.Id == pageId);
        if (page == null) throw new ArgumentOutOfRangeException(nameof(pageId), "The publication has no such page.");
        if (!page.SupportsPlacementEditing) throw new NotSupportedException("The page's coordinate registration is not qualified for text creation.");
        page.FrameEditRange.Validate(x, y, width, height);
        if (Text.Stories.Count >= options.Limits.MaxItems) throw new InvalidDataException("Publisher story limit exceeded.");
        if (NextObjectId == 0 || NextObjectId == uint.MaxValue)
            throw new InvalidOperationException("The publication has insufficient identifiers for a frame and its story.");
        frameId = NextObjectId; uint storyId = frameId + 1;
        PublisherTextStory story = PublisherStoryEditor.Create(storyId, paragraphs, options, token);
        if (Text.Stories.Values.Sum(PublisherStoryEditor.CharacterCount) + PublisherStoryEditor.CharacterCount(story) > options.Limits.MaxTextCharacters)
            throw new ArgumentException("Publication story text exceeds the configured character limit.", nameof(paragraphs));
        PublisherTextState text = Text.AddFrame(story, page, x, y, width, height, frameId);
        var owners = ObjectOwners.ToDictionary(item => item.Key, item => item.Value); owners.Add(frameId, pageId);
        uint createdId = frameId;
        return new PublisherPublicationState(Pages.Select(item => item.Id == pageId ? item.WithObject(createdId) : item),
            owners, Pictures, Images, text, TableStoryReferences, _sourceImageIds, NextImageId,
            storyId == uint.MaxValue ? 0 : storyId + 1, Shapes, Tables);
    }
}
