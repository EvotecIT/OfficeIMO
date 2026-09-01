using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Malformed_pages_body_uses_visual_fallback() {
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2)), new ulong[] { 2 }),
            ArchiveRecord(2, 2001, new byte[] { 0x80 }));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_BODY_MALFORMED");
    }

    [Fact]
    public void Malformed_table_models_use_visual_fallback() {
        byte[] records = Message(
            ArchiveRecord(1, 1,
                Message(ReferenceField(1, 2)), new ulong[] { 2 }),
            ArchiveRecord(2, 2,
                Message(StringField(1, "Sheet"), ReferenceField(2, 10)),
                new ulong[] { 10 }),
            ArchiveRecord(10, 6000,
                Message(ReferenceField(2, 11)), new ulong[] { 11 }),
            ArchiveRecord(11, 6001, new byte[] { 0x80 }));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_MODEL_UNSUPPORTED");
    }
}
