using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Image_metadata_catalog_is_independent_of_the_projected_image_limit() {
        using MemoryStream package = CreatePagesImagePackage(
            duplicateMetadata: false, imageCount: 1,
            imageBytes: ValidPreviewPng(), metadataEntryCount: 2);
        var options = new IWorkReadOptions {
            MaximumProjectedImages = 1,
            MaximumImageMetadataEntries = 2
        };
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages, options);
        IWorkArchiveRecord image = Assert.Single(source.Records,
            record => record.MessageType == 3005);

        IWorkImageAsset? asset = IWorkDrawingReader.ReadImage(source, image,
            new IWorkProjectionBudget(options), out bool complete);

        Assert.NotNull(asset);
        Assert.True(complete);
    }

    [Fact]
    public void Opened_directory_identity_rejects_a_replaced_bundle_root() {
        string parent = Path.Combine(Path.GetTempPath(), "officeimo-iwork-root-" + Guid.NewGuid().ToString("N"));
        string root = Path.Combine(parent, "Document.pages");
        string moved = Path.Combine(parent, "moved.pages");
        Directory.CreateDirectory(root);
        try {
            using var handle = OfficePathIdentity.OpenDirectoryForIdentity(root, out _);
            Directory.Move(root, moved);
            Directory.CreateDirectory(root);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                OfficePathIdentity.EnsurePathMatchesOpenedDirectory(root, handle));

            Assert.Contains("changed", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
            if (Directory.Exists(moved)) Directory.Delete(moved, recursive: true);
            if (Directory.Exists(parent)) Directory.Delete(parent, recursive: true);
        }
    }

    [Fact]
    public void Pages_text_boxes_without_geometry_use_visual_fallback() {
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2), ReferenceField(20, 3)), new ulong[] { 2, 3 }),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10020, Message(ReferenceField(1, 4)), new ulong[] { 4 }),
            ArchiveRecord(4, 2011,
                Message(BytesField(1, Message()), ReferenceField(2, 5)), new ulong[] { 5 }),
            ArchiveRecord(5, 2001, Message(StringField(3, "Floating"))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_text_boxes_without_effective_geometry_use_visual_fallback() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2)), new ulong[] { 2 }),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3))), new ulong[] { 3 }),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4)), new ulong[] { 4 }),
            ArchiveRecord(4, 5, Message(ReferenceField(7, 5)), new ulong[] { 5 }),
            ArchiveRecord(5, 2011,
                Message(BytesField(1, Message()), ReferenceField(2, 6)), new ulong[] { 6 }),
            ArchiveRecord(6, 2001, Message(StringField(3, "Floating"))));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED");
    }
}
