using System.Buffers.Binary;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Rejected_image_candidates_do_not_consume_the_shared_decode_budget() {
        byte[] rejected = CreateSizedPreviewPng(20, 20);
        rejected[^1] ^= 1;
        using MemoryStream package = CreatePagesImageCandidatesPackage(rejected,
            CreateSizedPreviewPng(20, 20));
        var options = new IWorkReadOptions { MaximumPackageBytes = 2500 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages, options);
        IWorkArchiveRecord image = Assert.Single(source.Records,
            record => record.MessageType == 3005);

        IWorkImageAsset? asset = IWorkDrawingReader.ReadImage(source, image,
            new IWorkProjectionBudget(options), out bool complete);

        Assert.NotNull(asset);
        Assert.True(complete);
        Assert.Equal("Data/valid.png", asset.PackagePath);
    }

    [Fact]
    public void Classic_zip_locator_signatures_inside_directory_records_are_not_zip64() {
        const int centralDirectorySize = 66;
        byte[] package = new byte[centralDirectorySize + 22];
        BinaryPrimitives.WriteUInt32LittleEndian(package.AsSpan(0, 4), 0x02014b50U);
        BinaryPrimitives.WriteUInt16LittleEndian(package.AsSpan(32, 2), 20);
        BinaryPrimitives.WriteUInt32LittleEndian(package.AsSpan(46, 4), 0x07064b50U);
        BinaryPrimitives.WriteUInt32LittleEndian(package.AsSpan(centralDirectorySize, 4),
            0x06054b50U);
        BinaryPrimitives.WriteUInt16LittleEndian(package.AsSpan(centralDirectorySize + 8, 2), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(package.AsSpan(centralDirectorySize + 10, 2), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(package.AsSpan(centralDirectorySize + 12, 4),
            centralDirectorySize);

        OfficeArchiveSafety.ZipCentralDirectoryScanResult byteResult =
            OfficeArchiveSafety.ScanZipCentralDirectory(package, 1);
        using var stream = new MemoryStream(package, writable: false);
        OfficeArchiveSafety.ZipCentralDirectoryScanResult streamResult =
            OfficeArchiveSafety.ScanZipCentralDirectory(stream, stream.Length, 1);

        Assert.True(byteResult.IsValid);
        Assert.Equal(1, byteResult.EntryCount);
        Assert.True(streamResult.IsValid);
        Assert.Equal(1, streamResult.EntryCount);
    }

    [Fact]
    public void Malformed_numbers_text_shapes_use_visual_fallback() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2)), new ulong[] { 2 }),
            ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 3)),
                new ulong[] { 3 }),
            ArchiveRecord(3, 2011, new byte[] { 0x80 }));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_TEXT_METADATA_UNSUPPORTED");
    }

    [Fact]
    public void Malformed_pages_image_records_use_visual_fallback() {
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2), ReferenceField(20, 3)), new ulong[] { 2, 3 }),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10020, Message(ReferenceField(1, 4)), new ulong[] { 4 }),
            ArchiveRecord(4, 3005, new byte[] { 0x80 }));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_IMAGE_UNSUPPORTED");
    }

    [Fact]
    public void Malformed_referenced_text_styles_use_visual_fallback() {
        byte[] styleTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, 3))));
        byte[] records = Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), new ulong[] { 2 }),
            ArchiveRecord(2, 2001,
                Message(StringField(3, "Styled"), BytesField(8, styleTable)), new ulong[] { 3 }),
            ArchiveRecord(3, 2021, new byte[] { 0x80 }));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Body.IsComplete);
    }

    [Fact]
    public void Malformed_pages_z_order_records_use_visual_fallback() {
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2), ReferenceField(20, 3)), new ulong[] { 2, 3 }),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10020, new byte[] { 0x80 }));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_shows_without_slide_sizes_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            omitSlideSize: true);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SLIDE_SIZE_UNSUPPORTED");
    }

    private static MemoryStream CreatePagesImageCandidatesPackage(byte[] rejected,
        byte[] valid) {
        byte[] image = Message(
            BytesField(1, Message()),
            BytesField(11, Message(VarintField(1, 100))),
            BytesField(15, Message(VarintField(1, 101))));
        byte[] metadata = Message(
            BytesField(4, Message(VarintField(1, 100),
                StringField(3, "rejected.png"), StringField(4, "rejected.png"))),
            BytesField(4, Message(VarintField(1, 101),
                StringField(3, "valid.png"), StringField(4, "valid.png"))));
        byte[] records = Message(
            ArchiveRecord(1, 10000, Message()),
            ArchiveRecord(2, 3005, image),
            ArchiveRecord(3, 11006, metadata));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("Data/rejected.png", rejected),
            ("Data/valid.png", valid));
    }
}
