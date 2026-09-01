using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Pages_owner_preserves_cross_type_drawable_stacking(bool imageFirst) {
        using MemoryStream package = CreatePagesPackageWithRestackedImageAndTextBox(imageFirst);

        using var result = WordIWorkConverter.LoadPagesWithReport(package);
        IWorkPagesDrawable[] drawables = result.Projection.Drawables.ToArray();
        WordTextBox textBox = Assert.Single(result.Document.TextBoxes);
        WordImage image = Assert.Single(result.Document.Images);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(imageFirst
                ? new[] { IWorkPagesDrawableKind.Image, IWorkPagesDrawableKind.TextBox }
                : new[] { IWorkPagesDrawableKind.TextBox, IWorkPagesDrawableKind.Image },
            drawables.Select(drawable => drawable.Kind));
        Assert.Equal(imageFirst, image.ZOrder < textBox.ZOrder);
        var mutable = Assert.IsAssignableFrom<IList<IWorkPagesDrawable>>(result.Projection.Drawables);
        Assert.Throws<NotSupportedException>(() => mutable[0] = mutable[0]);

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(imageFirst,
            Assert.Single(reopened.Images).ZOrder < Assert.Single(reopened.TextBoxes).ZOrder);
    }

    [Fact]
    public void Pages_visual_fallback_uses_the_recovered_page_layout() {
        byte[] layout = Message(
            FloatField(30, 842f), FloatField(31, 595f),
            FloatField(32, 54f), FloatField(33, 54f),
            FloatField(34, 36f), FloatField(35, 36f),
            FloatField(36, 18f), FloatField(37, 18f), VarintField(42, 1));
        using MemoryStream package = CreatePagesPackage(includeBody: false, textBox: null,
            includePreview: true, documentLayoutFields: layout);

        using var result = WordIWorkConverter.LoadPagesWithReport(package);
        WordSection section = Assert.Single(result.Document.Sections);
        WordImage preview = Assert.Single(result.Document.Images);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(16840U, section.PageSettings.Width.GetValueOrDefault());
        Assert.Equal(11900U, section.PageSettings.Height.GetValueOrDefault());
        Assert.Equal(1080U, section.Margins.Left);
        Assert.Equal(1080U, section.Margins.Right);
        Assert.InRange(preview.Width.GetValueOrDefault(), 1d, 734d);
        Assert.InRange(preview.Height.GetValueOrDefault(), 1d, 523d);

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(16840U, Assert.Single(reopened.Sections).PageSettings.Width.GetValueOrDefault());
    }

    [Fact]
    public void Keynote_visual_fallback_uses_the_recovered_slide_canvas() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            rotation: float.MaxValue, slideWidth: 720f, slideHeight: 360f);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);
        PowerPointPicture preview = Assert.Single(Assert.Single(result.Document.Slides).Pictures);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(720d, result.Document.SlideSize.WidthPoints, 3);
        Assert.Equal(360d, result.Document.SlideSize.HeightPoints, 3);
        Assert.Equal(5d, preview.WidthInches, 3);
        Assert.Equal(5d, preview.HeightInches, 3);
        Assert.Equal(2.5d, preview.LeftInches, 3);
        Assert.Equal(0d, preview.TopInches, 3);
    }

    private static MemoryStream CreatePagesPackageWithRestackedImageAndTextBox(bool imageFirst) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong zOrderId = 3;
        const ulong shapeId = 10;
        const ulong storageId = 11;
        const ulong imageId = 20;
        const ulong metadataId = 21;
        const ulong dataId = 22;
        const string imageName = "stacked.png";
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 36f), FloatField(2, 72f))),
            BytesField(2, Message(FloatField(1, 216f), FloatField(2, 108f))));
        byte[] drawable = Message(BytesField(1, geometry));
        byte[] shape = Message(BytesField(1, Message(BytesField(1, drawable))),
            ReferenceField(2, storageId));
        byte[] image = Message(BytesField(1, drawable),
            BytesField(11, Message(VarintField(1, dataId))));
        byte[] metadataEntry = Message(VarintField(1, dataId),
            StringField(3, imageName), StringField(4, imageName));
        byte[] zOrder = imageFirst
            ? Message(ReferenceField(1, imageId), ReferenceField(1, shapeId))
            : Message(ReferenceField(1, shapeId), ReferenceField(1, imageId));
        ulong[] zOrderReferences = imageFirst
            ? new[] { imageId, shapeId }
            : new[] { shapeId, imageId };
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId), ReferenceField(20, zOrderId)),
                new[] { bodyId, zOrderId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(zOrderId, 10020, zOrder, zOrderReferences),
            ArchiveRecord(shapeId, 2011, shape, new[] { storageId }),
            ArchiveRecord(storageId, 2001, Message(StringField(3, "Overlapping text"))),
            ArchiveRecord(imageId, 3005, image),
            ArchiveRecord(metadataId, 11006, Message(BytesField(4, metadataEntry))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ($"Data/{imageName}", ValidPreviewPng()),
            ("preview.png", ValidPreviewPng()));
    }
}
