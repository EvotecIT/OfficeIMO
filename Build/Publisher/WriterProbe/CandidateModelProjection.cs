using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects qualified source-free specimens from their final authoring models into codec inputs.</summary>
internal static class CandidateModelProjection {
    /// <summary>Projects the existing text-only two-centered-master or centered/left/right-master profiles.</summary>
    internal static CandidatePublicationSpec MultipleMasters(PublisherDocument document) {
        PublisherPage[] masters = document.MasterPages.ToArray();
        bool paired = masters.Length == 3 && masters[0].MasterSide == PublisherMasterPageSide.Single
            && masters[1].MasterSide == PublisherMasterPageSide.Left && masters[2].MasterSide == PublisherMasterPageSide.Right
            && masters[1].Name == masters[2].Name;
        bool centered = masters.Length == 2 && masters.All(master => master.MasterSide == PublisherMasterPageSide.Single);
        if (document.HasNativeSource || !paired && !centered || document.Pages.Count is < 1 or > 16
            || document.Pages.Concat(masters).Any(page => page.Width != 612 || page.Height != 792))
            throw new NotSupportedException("The master projection requires source-free Letter pages and the qualified master topology.");
        var indices = masters.Select((master, index) => (master.Id, index)).ToDictionary(item => item.Id, item => item.index);
        int?[] assignments = document.Pages.Select(page => page.MasterPageId is { } id && indices.TryGetValue(id, out int index)
            ? (int?)index : throw new NotSupportedException("The master specimen requires a defined master on every physical page.")).ToArray();
        var objects = new List<CandidateObjectSpec>();
        for (int index = 0; index < masters.Length; index++) AddPage(masters[index], index, true);
        for (int index = 0; index < document.Pages.Count; index++) AddPage(document.Pages[index], index, false);
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects, document.Pages.Count,
            masters.Select((master, index) => new CandidateMasterSpec(index == 0 ? "A" : paired ? "P" : "B",
                master.Name, master.MasterSide!.Value)).ToArray(), assignments);

        void AddPage(PublisherPage page, int index, bool onMaster) {
            if (page.Tables.Count != 0 || page.Pictures.Count != 0 || page.Shapes.Count != 0)
                throw new NotSupportedException("The multiple-master projection requires text-only page content.");
            var frames = page.TextFrames.ToDictionary(frame => frame.Id);
            foreach (uint id in page.ObjectIds) {
                if (!frames.TryGetValue(id, out PublisherTextFrame? frame) || frame.ColumnSpacing != 0)
                    throw new NotSupportedException("The multiple-master text specimen requires ordinary frames with the default column gap.");
                objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                    new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                    OrdinaryStory(document, frame), PageIndex: index, OnMaster: onMaster));
            }
        }
    }

    /// <summary>Projects ordinary single-master content, preserving each specimen's named or unnamed native picture profile.</summary>
    internal static CandidatePublicationSpec SingleMasterContent(PublisherDocument document, bool includePictureNames = true) {
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
                    objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                        new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                        OrdinaryStory(document, frame), PageIndex: index,
                        ColumnSpacing: frame.ColumnSpacing, OnMaster: onMaster));
                } else if (pictures.TryGetValue(id, out PublisherPicture? picture)) {
                    if (!picture.CanEdit || !picture.ImageId.HasValue)
                        throw new NotSupportedException("The picture has no qualified image and placement model.");
                    PublisherImage image = document.Images.Single(asset => asset.Id == picture.ImageId);
                    objects.Add(new CandidateObjectSpec(CandidateObjectKind.Picture,
                        new CandidatePlacement(picture.X, picture.Y, picture.Width, picture.Height),
                        Image: image.GetBytes(), PageIndex: index, Paint: picture.CreatePaint(),
                        ImageName: includePictureNames ? "Image" + image.Id + (image.ContentType == "image/png" ? ".png" : ".jpg") : null,
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

    private static CandidateStorySpec OrdinaryStory(PublisherDocument document, PublisherTextFrame frame) {
        if (frame.PreviousFrameId.HasValue || frame.NextFrameId.HasValue || frame.ColumnCount != 1
            || !frame.Padding.Equals(OfficeTextPadding.Empty) || frame.VerticalAlignment != OfficeTextVerticalAlignment.Top
            || frame.WrapObjectIds.Count != 0)
            throw new NotSupportedException("Linked, inset, multicolumn and wrapped frames require their separate writer profiles.");
        PublisherTextStory story = document.TextStories.Single(item => item.Id == frame.StoryId);
        if (story.Paragraphs.Count != 1)
            throw new NotSupportedException("The text specimen requires one complete paragraph per frame.");
        return CandidateStorySpec.Styled(story.Paragraphs[0]);
    }
}
