using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects qualified source-free specimens from their final authoring models into codec inputs.</summary>
internal static class CandidateModelProjection {
    /// <summary>Retains complete story identity and chain order independently of physical page and drawing order.</summary>
    internal static CandidatePublicationSpec LinkedStories(PublisherDocument document) {
        PublisherPage master = EmptySingleMaster(document);
        var frames = document.Pages.SelectMany((page, pageIndex) => FramePage(page)
            .Select(frame => (Frame: frame, PageIndex: pageIndex))).ToDictionary(item => item.Frame.Id);
        var ordered = document.Pages.Select(page => new Queue<PublisherTextFrame>(
            page.ObjectIds.Select(id => frames[id].Frame))).ToArray();
        var stories = new Dictionary<uint, CandidateStorySpec>();
        var objects = new List<CandidateObjectSpec>();
        // Keep the existing cross-page allocation slots; each page's final
        // object order supplies their current occupants.
        foreach (var slot in frames.Values.OrderBy(item => item.Frame.Id)) {
            PublisherTextFrame frame = ordered[slot.PageIndex].Dequeue();
            PlainFrameLayout(frame);
            if (!stories.TryGetValue(frame.StoryId, out CandidateStorySpec? story)) {
                story = CandidateStorySpec.Styled(document.TextStories.Single(item => item.Id == frame.StoryId));
                stories.Add(frame.StoryId, story);
            }
            objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height), story,
                PageIndex: slot.PageIndex, ColumnSpacing: frame.ColumnSpacing, FrameOrder: frame.Order));
        }
        return SingleMasterSpecification(document, master, objects);
    }

    /// <summary>Projects the qualified unlinked column, inset and Top/Center/Bottom frame profile.</summary>
    internal static CandidatePublicationSpec FrameLayouts(PublisherDocument document) {
        PublisherPage master = EmptySingleMaster(document);
        var objects = new List<CandidateObjectSpec>();
        for (int index = 0; index < document.Pages.Count; index++) {
            PublisherPage page = document.Pages[index];
            var frames = FramePage(page).ToDictionary(frame => frame.Id);
            foreach (uint id in page.ObjectIds) {
                PublisherTextFrame frame = frames[id];
                UnlinkedFrameLayout(frame);
                objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                    new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                    CandidateStorySpec.Styled(document.TextStories.Single(story => story.Id == frame.StoryId)),
                    PageIndex: index, FrameLayout: frame));
            }
        }
        return SingleMasterSpecification(document, master, objects);
    }

    /// <summary>Projects one complete unmerged borderless table and ordinary captions in final drawing order.</summary>
    internal static CandidatePublicationSpec Tables(PublisherDocument document) {
        PublisherPage master = EmptySingleMaster(document);
        if (document.Pages.Count != 1 || document.Pages[0].Tables.Count != 1
            || document.Pages[0].Pictures.Count != 0 || document.Pages[0].Shapes.Count != 0)
            throw new NotSupportedException("The table specimen requires one physical page and one borderless table with ordinary captions.");
        PublisherPage page = document.Pages[0];
        PublisherTable table = page.Tables[0];
        CompleteTable(table);
        var frames = page.TextFrames.ToDictionary(frame => frame.Id);
        var objects = new List<CandidateObjectSpec>();
        foreach (uint id in page.ObjectIds) {
            if (id == table.Id) {
                objects.Add(new CandidateObjectSpec(CandidateObjectKind.Table,
                    new CandidatePlacement(table.X, table.Y, table.Width, table.Height),
                    CandidateStorySpec.Table(table, document.TextStories.Single(story => story.Id == table.StoryId)), Table: table));
            } else if (frames.TryGetValue(id, out PublisherTextFrame? frame)) {
                objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                    new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                    OrdinaryStory(document, frame), ColumnSpacing: frame.ColumnSpacing));
            } else throw new NotSupportedException("An owned table-page object has no qualified projection.");
        }
        return SingleMasterSpecification(document, master, objects);
    }

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
    internal static CandidatePublicationSpec SingleMasterContent(PublisherDocument document, bool includePictureNames = true) =>
        Content(document, includePictureNames, composeProfiles: false);

    /// <summary>Combines ordinary master artwork with physical-page linked/layout text and one qualified table.</summary>
    internal static CandidatePublicationSpec ComposedContent(PublisherDocument document) =>
        Content(document, includePictureNames: true, composeProfiles: true);

    private static CandidatePublicationSpec Content(PublisherDocument document, bool includePictureNames, bool composeProfiles) {
        PublisherPage master = SingleMaster(document);
        if (composeProfiles && (master.Tables.Count != 0 || document.Pages.Sum(page => page.Tables.Count) > 1))
            throw new NotSupportedException("The composed specimen supports at most one physical-page table and no master table.");
        var objects = new List<CandidateObjectSpec>();
        var stories = new Dictionary<uint, CandidateStorySpec>();
        AddPage(master, 0, true);
        for (int index = 0; index < document.Pages.Count; index++) AddPage(document.Pages[index], index, false);
        return SingleMasterSpecification(document, master, objects);

        void AddPage(PublisherPage page, int index, bool onMaster) {
            if (!composeProfiles && page.Tables.Count != 0) throw new NotSupportedException("Tables require their separate writer profile.");
            var frames = page.TextFrames.ToDictionary(frame => frame.Id);
            var pictures = page.Pictures.ToDictionary(picture => picture.Id);
            var shapes = page.Shapes.ToDictionary(shape => shape.Id);
            var tables = page.Tables.ToDictionary(table => table.Id);
            foreach (uint id in page.ObjectIds) {
                if (frames.TryGetValue(id, out PublisherTextFrame? frame)) {
                    bool hasLayout = composeProfiles && !onMaster && !frame.PreviousFrameId.HasValue && !frame.NextFrameId.HasValue;
                    CandidateStorySpec story;
                    if (!composeProfiles || onMaster) story = OrdinaryStory(document, frame);
                    else {
                        if (hasLayout) UnlinkedFrameLayout(frame); else PlainFrameLayout(frame);
                        if (!stories.TryGetValue(frame.StoryId, out story!)) {
                            story = CandidateStorySpec.Styled(document.TextStories.Single(item => item.Id == frame.StoryId));
                            stories.Add(frame.StoryId, story);
                        }
                    }
                    objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox,
                        new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                        story, PageIndex: index, ColumnSpacing: frame.ColumnSpacing, OnMaster: onMaster,
                        FrameOrder: frame.Order, FrameLayout: hasLayout ? frame : null));
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
                } else if (tables.TryGetValue(id, out PublisherTable? table)) {
                    CompleteTable(table);
                    objects.Add(new CandidateObjectSpec(CandidateObjectKind.Table,
                        new CandidatePlacement(table.X, table.Y, table.Width, table.Height),
                        CandidateStorySpec.Table(table, document.TextStories.Single(story => story.Id == table.StoryId)),
                        PageIndex: index, Table: table));
                } else throw new NotSupportedException("An owned object has no supported model projection.");
            }
        }
    }

    private static CandidateStorySpec OrdinaryStory(PublisherDocument document, PublisherTextFrame frame) {
        if (frame.PreviousFrameId.HasValue || frame.NextFrameId.HasValue)
            throw new NotSupportedException("Linked frames require their separate writer profile.");
        PlainFrameLayout(frame);
        return CandidateStorySpec.Styled(document.TextStories.Single(item => item.Id == frame.StoryId));
    }

    private static void PlainFrameLayout(PublisherTextFrame frame) {
        if (frame.ColumnCount != 1 || !frame.Padding.Equals(OfficeTextPadding.Empty)
            || frame.VerticalAlignment != OfficeTextVerticalAlignment.Top || frame.WrapObjectIds.Count != 0)
            throw new NotSupportedException("Inset, multicolumn, aligned and wrapped frames require their separate writer profiles.");
    }

    private static void UnlinkedFrameLayout(PublisherTextFrame frame) {
        if (frame.PreviousFrameId.HasValue || frame.NextFrameId.HasValue || frame.WrapObjectIds.Count != 0
            || frame.ColumnCount is < 1 or > 3)
            throw new NotSupportedException("The layout specimen requires unlinked, unwrapped one-to-three-column frames.");
    }

    private static void CompleteTable(PublisherTable table) {
        if (table.RowCount < 1 || table.ColumnCount < 1 || table.Cells.Count > ushort.MaxValue
            || table.Cells.Sum(cell => (long)cell.RowSpan * cell.ColumnSpan) != (long)table.RowCount * table.ColumnCount || !table.HasTextMapping
            || !table.HasPlacement || !table.StoryId.HasValue)
            throw new NotSupportedException("The table specimen requires a complete grid within the native 16-bit cell count.");
    }

    private static IReadOnlyList<PublisherTextFrame> FramePage(PublisherPage page) {
        if (page.Tables.Count != 0 || page.Pictures.Count != 0 || page.Shapes.Count != 0
            || page.ObjectIds.Count != page.TextFrames.Count)
            throw new NotSupportedException("The frame specimen requires text-only page content.");
        return page.TextFrames;
    }

    private static PublisherPage EmptySingleMaster(PublisherDocument document) {
        PublisherPage master = SingleMaster(document);
        if (master.ObjectIds.Count != 0) throw new NotSupportedException("The text/table specimen requires an empty applied master.");
        return master;
    }

    private static PublisherPage SingleMaster(PublisherDocument document) {
        if (document.HasNativeSource || document.MasterPages.Count != 1
            || document.MasterPages[0].MasterSide != PublisherMasterPageSide.Single
            || document.Pages.Count is < 1 or > 16
            || document.Pages.Concat(document.MasterPages).Any(page => page.Width != 612 || page.Height != 792))
            throw new NotSupportedException("The model projection requires source-free Letter pages and one centered master.");
        PublisherPage master = document.MasterPages[0];
        if (document.Pages.Any(page => page.MasterPageId != master.Id))
            throw new NotSupportedException("The single-master specimen requires that master on every physical page.");
        return master;
    }

    private static CandidatePublicationSpec SingleMasterSpecification(PublisherDocument document, PublisherPage master,
        IReadOnlyList<CandidateObjectSpec> objects) => new(612 * 12700, 792 * 12700, objects, document.Pages.Count,
            new[] { new CandidateMasterSpec("A", master.Name, PublisherMasterPageSide.Single) },
            Enumerable.Repeat<int?>(0, document.Pages.Count).ToArray());
}
