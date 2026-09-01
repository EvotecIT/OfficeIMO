using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("A")]
    [InlineData("i")]
    public void Plain_alphabetic_keynote_numbering_uses_visual_fallback(string marker) {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(
            1, text: "Item", listLabel: marker);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Single(result.Document.Slides);
        Assert.Single(result.Document.Slides[0].Pictures);
        Assert.Contains(result.ImportReport.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_POWERPOINT_DESTINATION_UNSUPPORTED");
    }

    [Fact]
    public void Conflicting_keynote_text_storage_references_disable_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithStorageReferences(
            field2StorageId: 6, field4StorageId: 7);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Aliased_keynote_text_storage_references_remain_editable() {
        using MemoryStream package = CreateKeynotePackageWithStorageReferences(
            field2StorageId: 6, field4StorageId: 6);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("Primary", Assert.Single(
            Assert.Single(result.Projection.Slides).TextBoxes).Content.PlainText);
    }

    [Fact]
    public void Duplicate_keynote_text_storage_references_disable_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithStorageReferences(
            field2StorageId: 6, field4StorageId: 6, duplicateField2: true);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Sub_tick_numbers_durations_preserve_their_excel_day_fraction(bool formula) {
        const double seconds = 0.00000001d;
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Sub-tick duration", 1, 1, seconds,
                hasFormula: formula, duration: true)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        double expected = seconds / 86_400d;

        Assert.False(result.IsVisualFallback);
        Assert.Equal(expected, result.Document.Sheets[0].CellAt(1, 1).GetValue<double>());

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(5)]
    [InlineData(6)]
    public void Ordered_keynote_fields_determine_placeholder_stacking(int placeholderField) {
        using MemoryStream package = CreateKeynotePackageWithOrderedPlaceholder(placeholderField);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        IWorkKeynoteSlide sourceSlide = Assert.Single(result.Projection.Slides);
        PowerPointSlide targetSlide = Assert.Single(result.Document.Slides);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(new[] { IWorkKeynoteDrawableKind.Image, IWorkKeynoteDrawableKind.TextBox },
            sourceSlide.Drawables.Select(drawable => drawable.Kind));
        Assert.IsType<PowerPointPicture>(targetSlide.Shapes[0]);
        Assert.IsType<PowerPointTextBox>(targetSlide.Shapes[1]);
    }

    private static MemoryStream CreateKeynotePackageWithStorageReferences(
        ulong field2StorageId, ulong field4StorageId, bool duplicateField2 = false) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong primaryStorageId = 6;
        const ulong alternateStorageId = 7;
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(7, shapeId))),
            ArchiveRecord(shapeId, 2011,
                Message(ReferenceField(2, field2StorageId),
                    duplicateField2 ? ReferenceField(2, field2StorageId) : Array.Empty<byte>(),
                    ReferenceField(4, field4StorageId)),
                new[] { field2StorageId, field4StorageId }),
            ArchiveRecord(primaryStorageId, 2001, Message(StringField(3, "Primary"))),
            ArchiveRecord(alternateStorageId, 2001, Message(StringField(3, "Alternate"))));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreateKeynotePackageWithOrderedPlaceholder(int placeholderField) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        const ulong imageId = 10;
        const ulong metadataId = 13;
        const ulong dataId = 20;
        const string imageName = "ordered-placeholder.png";
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 72f), FloatField(2, 72f))),
            BytesField(2, Message(FloatField(1, 144f), FloatField(2, 96f))));
        byte[] image = Message(
            BytesField(1, Message(BytesField(1, geometry))),
            BytesField(11, Message(VarintField(1, dataId))));
        byte[] metadata = Message(BytesField(4, Message(VarintField(1, dataId),
            StringField(3, imageName), StringField(4, imageName))));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5,
                Message(ReferenceField(placeholderField, shapeId),
                    ReferenceField(7, imageId), ReferenceField(7, shapeId)),
                new[] { shapeId, imageId }),
            ArchiveRecord(shapeId, 2011,
                Message(ReferenceField(2, storageId)), new[] { storageId }),
            ArchiveRecord(storageId, 2001, Message(StringField(3, "Placeholder"))),
            ArchiveRecord(imageId, 3005, image),
            ArchiveRecord(metadataId, 11006, metadata));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ($"Data/{imageName}", ValidPreviewPng()),
            ("preview.png", ValidPreviewPng()));
    }
}
