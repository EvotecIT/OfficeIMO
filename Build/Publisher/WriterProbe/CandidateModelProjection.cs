using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects the existing single-master specimen from its final authoring model into codec inputs.</summary>
internal static class CandidateModelProjection {
    internal static CandidatePublicationSpec SingleMasterContent(PublisherDocument document) {
        if (document.HasNativeSource || document.MasterPages.Count != 1
            || document.MasterPages[0].MasterSide != PublisherMasterPageSide.Single
            || document.Pages.Count is < 1 or > 16
            || document.Pages.Concat(document.MasterPages).Any(page => page.Width != 612 || page.Height != 792))
            throw new NotSupportedException("The model projection requires source-free Letter pages and one centered master.");
        PublisherPage master = document.MasterPages[0];
        if (document.Pages.Any(page => page.MasterPageId != master.Id))
            throw new NotSupportedException("The single-master specimen requires that master on every physical page.");
        var objects = new List<CandidateObjectSpec>();
        AddPage(master, 0, true);
        for (int index = 0; index < document.Pages.Count; index++) AddPage(document.Pages[index], index, false);
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects, document.Pages.Count,
            new[] { new CandidateMasterSpec("A", master.Name, PublisherMasterPageSide.Single) },
            Enumerable.Repeat<int?>(0, document.Pages.Count).ToArray());

        void AddPage(PublisherPage page, int index, bool onMaster) {
            if (page.Tables.Count != 0) throw new NotSupportedException("Tables require their separate writer profile.");
            var frames = page.TextFrames.ToDictionary(frame => frame.Id);
            var pictures = page.Pictures.ToDictionary(picture => picture.Id);
            var shapes = page.Shapes.ToDictionary(shape => shape.Id);
            foreach (uint id in page.ObjectIds) {
                if (frames.TryGetValue(id, out PublisherTextFrame? frame)) {
                    if (frame.PreviousFrameId.HasValue || frame.NextFrameId.HasValue || frame.ColumnCount != 1
                        || !frame.Padding.Equals(OfficeTextPadding.Empty) || frame.VerticalAlignment != OfficeTextVerticalAlignment.Top
                        || frame.WrapObjectIds.Count != 0)
                        throw new NotSupportedException("Linked, inset, multicolumn and wrapped frames require their separate writer profiles.");
                    PublisherTextStory story = document.TextStories.Single(item => item.Id == frame.StoryId);
                    if (story.Paragraphs.Count != 1)
                        throw new NotSupportedException("The single-master text specimen requires one complete paragraph per frame.");
                    objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                        new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                        CandidateStorySpec.Styled(story.Paragraphs[0]), PageIndex: index,
                        ColumnSpacing: frame.ColumnSpacing, OnMaster: onMaster));
                } else if (pictures.TryGetValue(id, out PublisherPicture? picture)) {
                    if (!picture.CanEdit || !picture.ImageId.HasValue)
                        throw new NotSupportedException("The picture has no qualified image and placement model.");
                    PublisherImage image = document.Images.Single(asset => asset.Id == picture.ImageId);
                    objects.Add(new CandidateObjectSpec(CandidateObjectKind.Picture,
                        new CandidatePlacement(picture.X, picture.Y, picture.Width, picture.Height),
                        Image: image.GetBytes(), PageIndex: index, Paint: picture.CreatePaint(),
                        ImageName: "Image" + image.Id + (image.ContentType == "image/png" ? ".png" : ".jpg"),
                        Crop: picture.Crop, ImageId: image.Id, OnMaster: onMaster));
                } else if (shapes.TryGetValue(id, out PublisherShape? shape)) {
                    OfficeShape paint = shape.CreateShape();
                    if (paint.Kind != OfficeShapeKind.Rectangle)
                        throw new NotSupportedException("The single-master artwork specimen requires rectangles.");
                    objects.Add(new CandidateObjectSpec(CandidateObjectKind.Rectangle,
                        new CandidatePlacement(shape.X, shape.Y, shape.Width, shape.Height),
                        PageIndex: index, Paint: paint, OnMaster: onMaster));
                } else throw new NotSupportedException("An owned object has no supported model projection.");
            }
        }
    }
}
