using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_documents_without_supported_sheets_use_visual_fallback() {
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(
                ArchiveRecord(1, 1, Message()))),
            ("preview.png", ValidPreviewPng()));

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_SHEET_MISSING");
    }

    [Fact]
    public void Formula_flags_on_empty_numbers_cells_use_visual_fallback() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Empty formula", 1, 1, 0d, hasFormula: true,
                completeFormula: true, emptyCellWithFormula: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_CELL_DECODE");
    }

    [Fact]
    public void Every_protobuf_entry_point_rejects_field_numbers_above_29_bits() {
        byte[] invalid = Message(Varint(1UL << 32), Varint(0));
        var options = new IWorkReadOptions();

        Assert.Throws<InvalidDataException>(() => IWorkProtobuf.Parse(invalid, options));
        Assert.Throws<InvalidDataException>(() => IWorkProtobuf.CountFields(
            invalid, targetField: 1, maximumFields: 10));
        Assert.Throws<InvalidDataException>(() => IWorkProtobuf.ParseRepeatedMessages(
            invalid, targetField: 1, maximumMatches: 10, options, out _));
    }

    [Fact]
    public void Classic_pdf_xrefs_reject_reserved_or_out_of_range_in_use_generations() {
        Assert.False(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(
            validKids: true, generation: 65535)));
        Assert.False(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(
            validKids: true, generation: 99999)));
    }

    [Fact]
    public void Empty_pages_floating_drawable_entries_use_visual_fallback() {
        byte[] pageGroup = Message(BytesField(2, Message()));
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2), ReferenceField(3, 3)),
                new ulong[] { 2, 3 }),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10020, Message(BytesField(1, pageGroup))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }
}
