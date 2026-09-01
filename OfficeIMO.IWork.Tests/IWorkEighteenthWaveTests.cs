using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Build_history_rejects_nested_string_elements_without_materializing_descendant_values() {
        const int depth = 128;
        string nested = string.Concat(Enumerable.Repeat("<string>", depth))
            + "version"
            + string.Concat(Enumerable.Repeat("</string>", depth));
        string plist = "<?xml version=\"1.0\"?><plist version=\"1.0\"><array>"
            + nested + "</array></plist>";
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(ArchiveRecord(1, 1, Message()))),
            ("Metadata/BuildVersionHistory.plist", Encoding.UTF8.GetBytes(plist)));

        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Numbers);

        Assert.Empty(source.BuildVersions);
    }

    [Fact]
    public void Numbers_row_heights_outside_the_xlsx_range_use_visual_fallback() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Tall", 1, 1, 42d, defaultRowHeight: 410d)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Numbers_column_widths_outside_the_xlsx_range_use_visual_fallback() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Wide", 1, 1, 42d, defaultColumnWidth: 1786d)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Png_validation_accepts_empty_idat_chunks_when_the_aggregate_stream_is_valid() {
        byte[] png = InsertEmptyIdatChunk(ValidPreviewPng());

        (int? width, int? height) = IWorkImageInfo.Read(
            png, "image/png", 16 * 1024 * 1024);

        Assert.True(width > 0);
        Assert.True(height > 0);
    }

    [Fact]
    public void Reachable_pages_shapes_without_text_storage_use_visual_fallback() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong shapeId = 3;
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)),
                new[] { bodyId, shapeId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(shapeId, 2011, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    private static byte[] InsertEmptyIdatChunk(byte[] png) {
        int offset = 8;
        while (offset <= png.Length - 12) {
            int length = png[offset] << 24 | png[offset + 1] << 16
                | png[offset + 2] << 8 | png[offset + 3];
            if (png[offset + 4] == (byte)'I' && png[offset + 5] == (byte)'D'
                && png[offset + 6] == (byte)'A' && png[offset + 7] == (byte)'T') {
                return Message(png.AsSpan(0, offset).ToArray(),
                    CreatePngChunk("IDAT", Array.Empty<byte>()),
                    png.AsSpan(offset).ToArray());
            }
            offset = checked(offset + 12 + length);
        }
        throw new InvalidDataException("The test PNG has no IDAT chunk.");
    }
}
