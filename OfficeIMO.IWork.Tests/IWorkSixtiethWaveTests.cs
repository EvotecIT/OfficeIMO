using System.Buffers.Binary;
using OfficeIMO.Core.Internal;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Zip_entry_limits_are_enforced_before_archive_metadata_is_materialized(bool nested) {
        using MemoryStream inner = CreatePackage(
            ("Index/Document.iwa", FrameIwa(ArchiveRecord(1, 1, Array.Empty<byte>()))));
        using MemoryStream oversizedDirectory = PatchDeclaredZipEntryCount(inner, 4096);
        using MemoryStream package = nested
            ? CreatePackage(("Index.zip", oversizedDirectory.ToArray()),
                ("preview.png", ValidPreviewPng()))
            : new MemoryStream(oversizedDirectory.ToArray(), writable: false);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumEntryCount = nested ? 2 : 1 }));

        Assert.Contains("before package entries are opened", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Keynote_slide_sizes_below_emu_precision_report_editable_quantization() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            slideWidth: 960.0001f, slideHeight: 540f);

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_PPTX_PRECISION");
    }

    [Fact]
    public void Keynote_drawable_geometry_below_emu_precision_reports_editable_quantization() {
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 10.0001f), FloatField(2, 10f))),
            BytesField(2, Message(FloatField(1, 40f), FloatField(2, 30f))));
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            textBoxDrawable: Message(BytesField(1, geometry)));

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_PPTX_PRECISION");
    }

    [Fact]
    public void Keynote_table_rows_below_emu_precision_report_editable_quantization() {
        using MemoryStream package = CreateKeynotePackageWithTableDefaults(
            rows: 1, columns: 1,
            defaultRowHeight: 0.5d / PowerPointUnits.EmusPerPoint,
            defaultColumnWidth: 30d, includePreview: true);

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_PPTX_PRECISION");
        PowerPointTable table = Assert.Single(result.Value.Slides[0].Tables);
        Assert.True(table.Height > 0);
        Assert.True(table.GetRowHeight(0) > 0);
    }

    [Fact]
    public void Keynote_image_extents_below_half_an_emu_remain_editable() {
        using MemoryStream package = CreateKeynotePackageWithTinyImageExtent();

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_PPTX_PRECISION");
        PowerPointPicture picture = Assert.Single(result.Value.Slides[0].Pictures);
        Assert.True(picture.Width > 0);
        Assert.True(picture.Height > 0);
    }

    [Fact]
    public void Classic_zip_can_declare_exactly_65535_entries_without_zip64() {
        byte[] package = CreateClassicCentralDirectory(ushort.MaxValue);

        OfficeArchiveSafety.ZipCentralDirectoryScanResult byteResult =
            OfficeArchiveSafety.ScanZipCentralDirectory(package, ushort.MaxValue);
        using var stream = new MemoryStream(package, writable: false);
        OfficeArchiveSafety.ZipCentralDirectoryScanResult streamResult =
            OfficeArchiveSafety.ScanZipCentralDirectory(stream, stream.Length,
                ushort.MaxValue);

        Assert.True(byteResult.IsValid);
        Assert.False(byteResult.LimitExceeded);
        Assert.Equal(ushort.MaxValue, byteResult.EntryCount);
        Assert.True(streamResult.IsValid);
        Assert.False(streamResult.LimitExceeded);
        Assert.Equal(ushort.MaxValue, streamResult.EntryCount);
    }

    [Fact]
    public void Zip64_central_directory_is_validated_by_both_bounded_scanners() {
        byte[] package = CreateZip64CentralDirectory(malformedRecordOffset: false);

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
    public void Malformed_zip64_record_offsets_are_rejected_by_both_bounded_scanners() {
        byte[] package = CreateZip64CentralDirectory(malformedRecordOffset: true);

        OfficeArchiveSafety.ZipCentralDirectoryScanResult byteResult =
            OfficeArchiveSafety.ScanZipCentralDirectory(package, 1);
        using var stream = new MemoryStream(package, writable: false);
        OfficeArchiveSafety.ZipCentralDirectoryScanResult streamResult =
            OfficeArchiveSafety.ScanZipCentralDirectory(stream, stream.Length, 1);

        Assert.False(byteResult.IsValid);
        Assert.Contains("declared offset", byteResult.Error,
            StringComparison.OrdinalIgnoreCase);
        Assert.False(streamResult.IsValid);
        Assert.Contains("declared offset", streamResult.Error,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Low_count_central_directory_mismatches_are_rejected_before_opening() {
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(ArchiveRecord(1, 1, Array.Empty<byte>()))));
        using MemoryStream malformed = PatchDeclaredZipEntryCount(package, 2);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(malformed, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumEntryCount = 10 }));

        Assert.Contains("declares 2 entries but contains 1", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private static MemoryStream CreateKeynotePackageWithTinyImageExtent() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong imageId = 5;
        const ulong metadataId = 6;
        const ulong dataId = 7;
        const string imageName = "tiny.png";
        float tinyPoints = (float)(0.49d / PowerPointUnits.EmusPerPoint);
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 10f), FloatField(2, 10f))),
            BytesField(2, Message(FloatField(1, tinyPoints), FloatField(2, tinyPoints))));
        byte[] image = Message(
            BytesField(1, Message(BytesField(1, geometry))),
            BytesField(11, Message(VarintField(1, dataId))));
        byte[] metadata = Message(BytesField(4, Message(
            VarintField(1, dataId), StringField(3, imageName), StringField(4, imageName))));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(7, imageId))),
            ArchiveRecord(imageId, 3005, image),
            ArchiveRecord(metadataId, 11006, metadata));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ($"Data/{imageName}", ValidPreviewPng()),
            ("preview.png", ValidPreviewPng()));
    }

    private static byte[] CreateClassicCentralDirectory(ushort entryCount) {
        const int centralHeaderLength = 46;
        int centralDirectorySize = checked(entryCount * centralHeaderLength);
        byte[] bytes = new byte[checked(centralDirectorySize + 22)];
        for (int offset = 0; offset < centralDirectorySize; offset += centralHeaderLength) {
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset, 4), 0x02014b50U);
        }
        int end = centralDirectorySize;
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(end, 4), 0x06054b50U);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(end + 8, 2), entryCount);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(end + 10, 2), entryCount);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(end + 12, 4),
            checked((uint)centralDirectorySize));
        return bytes;
    }

    private static byte[] CreateZip64CentralDirectory(bool malformedRecordOffset) {
        const int centralDirectorySize = 46;
        const int zip64RecordOffset = centralDirectorySize;
        const int zip64LocatorOffset = zip64RecordOffset + 56;
        const int endOffset = zip64LocatorOffset + 20;
        byte[] bytes = new byte[endOffset + 22];
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(0, 4), 0x02014b50U);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(zip64RecordOffset, 4),
            0x06064b50U);
        BinaryPrimitives.WriteUInt64LittleEndian(bytes.AsSpan(zip64RecordOffset + 4, 8), 44);
        BinaryPrimitives.WriteUInt64LittleEndian(bytes.AsSpan(zip64RecordOffset + 24, 8), 1);
        BinaryPrimitives.WriteUInt64LittleEndian(bytes.AsSpan(zip64RecordOffset + 32, 8), 1);
        BinaryPrimitives.WriteUInt64LittleEndian(bytes.AsSpan(zip64RecordOffset + 40, 8),
            centralDirectorySize);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(zip64LocatorOffset, 4),
            0x07064b50U);
        BinaryPrimitives.WriteUInt64LittleEndian(bytes.AsSpan(zip64LocatorOffset + 8, 8),
            malformedRecordOffset ? (ulong)(zip64RecordOffset - 1) : zip64RecordOffset);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(zip64LocatorOffset + 16, 4), 1);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(endOffset, 4), 0x06054b50U);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(endOffset + 8, 2),
            ushort.MaxValue);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(endOffset + 10, 2),
            ushort.MaxValue);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(endOffset + 12, 4),
            centralDirectorySize);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(endOffset + 16, 4),
            uint.MaxValue);
        return bytes;
    }

    private static MemoryStream PatchDeclaredZipEntryCount(MemoryStream package,
        ushort count) {
        byte[] bytes = package.ToArray();
        for (int offset = bytes.Length - 22; offset >= 0; offset--) {
            if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4))
                    != 0x06054b50U) continue;
            ushort commentLength = BinaryPrimitives.ReadUInt16LittleEndian(
                bytes.AsSpan(offset + 20, 2));
            if (offset + 22 + commentLength != bytes.Length) continue;
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset + 8, 2), count);
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset + 10, 2), count);
            return new MemoryStream(bytes, writable: false);
        }
        throw new InvalidOperationException("The test package has no ZIP end record.");
    }
}
