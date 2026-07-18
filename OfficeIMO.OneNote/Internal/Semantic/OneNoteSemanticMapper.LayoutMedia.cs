namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    private static OneNoteImage BuildImage(
        OneNoteMaterializedObjectSpace space,
        OneNoteRevisionStoreObject item,
        OneNoteObjectSpaceMaterializer materializer) {
        var image = new OneNoteImage {
            Id = item.Id,
            FileName = ReadString(item, ImageFilename),
            AltText = ReadString(item, ImageAltText),
            SourcePath = ReadString(item, SourceFilePath),
            Hyperlink = ReadString(item, HyperlinkUrl),
            OcrText = ReadString(item, OneNoteSchema.RichEditTextUnicode),
            OcrLanguageId = ReadUInt32(item, OneNoteSchema.LanguageId),
            DisplayedPageNumber = ReadUInt32(item, OneNoteSchema.DisplayedPageNumber),
            IsBackground = ReadBoolean(item, OneNoteSchema.IsBackground),
            SizeSetByUser = ReadBoolean(item, OneNoteSchema.IsLayoutSizeSetByUser),
            UploadState = ReadUInt32(item, OneNoteSchema.ImageUploadState),
            WidthHalfInches = ReadFloat(item, PictureWidth),
            HeightHalfInches = ReadFloat(item, PictureHeight),
            Layout = ReadLayout(item)
        };
        image.MediaType = ResolveMediaType(image.FileName);
        image.PictureContainerObjectId = GetReferences(item, PictureContainer).FirstOrDefault();
        image.WebPictureContainerObjectId = GetReferences(item, OneNoteSchema.WebPictureContainer14).FirstOrDefault();
        OneNoteRevisionStoreObject? picture = image.PictureContainerObjectId == null
            ? null
            : space.GetObject(image.PictureContainerObjectId);
        OneNoteRevisionStoreObject? webPicture = image.WebPictureContainerObjectId == null
            ? null
            : space.GetObject(image.WebPictureContainerObjectId);
        PopulateBinaryPayload(image, picture, materializer);
        if (image.Payload == null && webPicture != null) {
            PopulateBinaryPayload(image, webPicture, materializer);
            image.PayloadUsesWebPictureContainer = image.Payload != null;
        }
        ApplyTags(image, item, space);
        return image;
    }

    private static OneNoteLayout ReadLayout(OneNoteRevisionStoreObject item) {
        return new OneNoteLayout {
            X = ReadFloat(item, OffsetFromParentHorizontal),
            Y = ReadFloat(item, OffsetFromParentVertical),
            Width = ReadFloat(item, LayoutMaxWidth),
            Height = ReadFloat(item, LayoutMaxHeight),
            Tight = ReadBoolean(item, LayoutTightLayout),
            RightToLeft = ReadBoolean(item, OutlineElementRtl),
            MinimumWidth = ReadFloat(item, OneNoteSchema.LayoutMinimumOutlineWidth),
            AlignmentInParent = ReadUInt32(item, OneNoteSchema.LayoutAlignmentInParent),
            AlignmentSelf = ReadUInt32(item, OneNoteSchema.LayoutAlignmentSelf),
            CollisionPriority = ReadUInt32(item, OneNoteSchema.LayoutCollisionPriority),
            TightAlignment = ReadBoolean(item, OneNoteSchema.LayoutTightAlignment)
        };
    }

    private static OneNotePageSize? ReadPageSize(OneNoteRevisionStoreObject pageNode) {
        uint? value = ReadUInt32(pageNode, OneNoteSchema.PageSize);
        return value.HasValue && value.Value <= (uint)OneNotePageSize.Custom ? (OneNotePageSize)value.Value : null;
    }

    private static OneNotePageOrientation? ReadPageOrientation(OneNoteRevisionStoreObject pageNode) {
        bool? portrait = ReadBoolean(pageNode, OneNoteSchema.PortraitPage);
        return portrait.HasValue
            ? portrait.Value ? OneNotePageOrientation.Portrait : OneNotePageOrientation.Landscape
            : (OneNotePageOrientation?)null;
    }
}
