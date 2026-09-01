using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Geometry_free_keynote_tables_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithTableDefaults(
            rows: 2, columns: 2, defaultRowHeight: 20d, defaultColumnWidth: 40d,
            includePreview: true, omitTableGeometry: true);

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_DRAWABLE_UNSUPPORTED");
        Assert.Single(Assert.Single(result.Value.Slides).Pictures);
    }

    [Fact]
    public void Malformed_pages_image_catalogs_use_visual_fallback() {
        using MemoryStream package = CreatePackageWithMalformedImageCatalog(
            IWorkDocumentKind.Pages);

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_IMAGE_UNSUPPORTED");
    }

    [Fact]
    public void Malformed_keynote_image_catalogs_use_visual_fallback() {
        using MemoryStream package = CreatePackageWithMalformedImageCatalog(
            IWorkDocumentKind.Keynote);

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_IMAGE_UNSUPPORTED");
    }

    private static MemoryStream CreatePackageWithMalformedImageCatalog(IWorkDocumentKind kind) {
        const ulong imageId = 5;
        const ulong dataId = 20;
        const string imageName = "catalog.png";
        byte[] image = Message(
            BytesField(1, Message(GeometryDrawable(10f, 10f, 100f, 50f))),
            BytesField(11, Message(VarintField(1, dataId))));
        byte[] graph = kind == IWorkDocumentKind.Pages
            ? Message(
                ArchiveRecord(1, 10000,
                    Message(ReferenceField(4, 2), ReferenceField(20, 3)),
                    new ulong[] { 2, 3 }),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
                ArchiveRecord(3, 10020, Message(ReferenceField(1, imageId)),
                    new ulong[] { imageId }),
                ArchiveRecord(imageId, 3005, image),
                ArchiveRecord(6, 11006, new byte[] { 0x80 }))
            : Message(
                ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(7, imageId))),
                ArchiveRecord(imageId, 3005, image),
                ArchiveRecord(6, 11006, new byte[] { 0x80 }));
        return CreatePackage(
            (kind == IWorkDocumentKind.Pages ? "Index/Document.iwa" : "Index/Slide.iwa",
                FrameIwa(graph)),
            ($"Data/{imageName}", ValidPreviewPng()),
            ("preview.png", ValidPreviewPng()));
    }
}
