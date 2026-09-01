using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_visual_only_mode_ignores_the_semantic_sheet_limit() {
        using MemoryStream package = CreateNumbersPackage(
            Array.Empty<TableSpec>(), includePreview: true, sheetReferenceCount: 2);

        using var result = ExcelDocument.LoadNumbersWithReport(package,
            new IWorkReadOptions {
                ImportMode = IWorkImportMode.VisualOnly,
                MaximumProjectedSheets = 1
            });

        Assert.True(result.IsVisualFallback);
        Assert.Empty(result.Projection.Sheets);
        Assert.NotNull(result.ImportReport.VisualPreview);
    }

    [Fact]
    public void Keynote_repeated_image_uses_are_bounded_by_encoded_destination_bytes() {
        byte[] image = CreatePaddedValidPng();
        using MemoryStream package = CreateKeynotePackageWithSharedImage(image);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            PowerPointPresentation.LoadKeynoteWithReport(package,
                new IWorkReadOptions { MaximumProjectedImageBytes = image.LongLength }));

        Assert.Contains("destination image", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Pages_repeated_image_uses_share_the_destination_byte_budget() {
        byte[] image = CreatePaddedValidPng();
        using MemoryStream package = CreatePagesPackageWithSharedImage(image);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            WordDocument.LoadPagesWithReport(package,
                new IWorkReadOptions { MaximumProjectedImageBytes = image.LongLength }));

        Assert.Contains("destination image", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] CreatePaddedValidPng() {
        byte[] source = ValidPreviewPng();
        byte[] text = System.Text.Encoding.ASCII.GetBytes(
            "Comment\0" + new string('x', 4088));
        return Message(source[..^12],
            CreatePngChunk("tEXt", text), source[^12..]);
    }

    private static MemoryStream CreateKeynotePackageWithSharedImage(byte[] imageBytes) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong firstNodeId = 3;
        const ulong secondNodeId = 4;
        const ulong firstSlideId = 5;
        const ulong secondSlideId = 6;
        const ulong imageId = 7;
        const ulong metadataId = 8;
        const ulong dataId = 9;
        const string imageName = "shared.png";
        byte[] slideTree = Message(
            ReferenceField(2, firstNodeId), ReferenceField(2, secondNodeId));
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 0f), FloatField(2, 0f))),
            BytesField(2, Message(FloatField(1, 64f), FloatField(2, 64f))));
        byte[] image = Message(
            BytesField(1, Message(BytesField(1, geometry))),
            BytesField(11, Message(VarintField(1, dataId))));
        byte[] metadata = Message(BytesField(4, Message(
            VarintField(1, dataId), StringField(3, imageName), StringField(4, imageName))));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId)),
                new[] { showId }),
            ArchiveRecord(showId, 2, Message(BytesField(3, slideTree)),
                new[] { firstNodeId, secondNodeId }),
            ArchiveRecord(firstNodeId, 4, Message(ReferenceField(2, firstSlideId)),
                new[] { firstSlideId }),
            ArchiveRecord(secondNodeId, 4, Message(ReferenceField(2, secondSlideId)),
                new[] { secondSlideId }),
            ArchiveRecord(firstSlideId, 5, Message(ReferenceField(7, imageId)),
                new[] { imageId }),
            ArchiveRecord(secondSlideId, 5, Message(ReferenceField(7, imageId)),
                new[] { imageId }),
            ArchiveRecord(imageId, 3005, image),
            ArchiveRecord(metadataId, 11006, metadata));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ($"Data/{imageName}", imageBytes),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreatePagesPackageWithSharedImage(byte[] imageBytes) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstImageId = 3;
        const ulong secondImageId = 4;
        const ulong metadataId = 5;
        const ulong dataId = 6;
        const string imageName = "shared.png";
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 0f), FloatField(2, 0f))),
            BytesField(2, Message(FloatField(1, 64f), FloatField(2, 64f))));
        byte[] image = Message(
            BytesField(1, Message(BytesField(1, geometry))),
            BytesField(11, Message(VarintField(1, dataId))));
        byte[] metadata = Message(BytesField(4, Message(
            VarintField(1, dataId), StringField(3, imageName), StringField(4, imageName))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)),
                new[] { bodyId, firstImageId, secondImageId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(firstImageId, 3005, image),
            ArchiveRecord(secondImageId, 3005, image),
            ArchiveRecord(metadataId, 11006, metadata));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ($"Data/{imageName}", imageBytes),
            ("preview.png", ValidPreviewPng()));
    }
}
