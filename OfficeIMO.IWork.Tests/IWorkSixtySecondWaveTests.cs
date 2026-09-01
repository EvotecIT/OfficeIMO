using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    public void Pdf_dictionaries_require_their_own_object_terminator(
        bool omitCatalogEndObject, bool omitPagesEndObject, bool omitPageEndObject) {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            omitCatalogEndObject: omitCatalogEndObject,
            omitPagesEndObject: omitPagesEndObject,
            omitPageEndObject: omitPageEndObject);

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    [Fact]
    public void Conflicting_text_clear_and_value_directives_disable_editable_reconstruction() {
        for (int kind = 0; kind < 3; kind++) {
            using MemoryStream package = CreatePagesPackageWithStyleChain(depth: 1,
                includePreview: true, conflictingFont: kind == 0,
                conflictingColor: kind == 1, conflictingBackground: kind == 2);

            using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

            Assert.True(result.IsVisualFallback);
            Assert.False(result.Projection.Body.IsComplete);
            Assert.Contains(result.Projection.Diagnostics, diagnostic =>
                diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
        }
    }

    [Fact]
    public void Png_zlib_streams_reject_bytes_after_the_deflate_terminator() {
        byte[] png = CreatePngWithTrailingDeflateByte();

        (int? width, int? height) = IWorkImageInfo.Read(png, "image/png",
            maximumDecodedBytes: 1024);

        Assert.Null(width);
        Assert.Null(height);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    public void Zero_dimensional_numbers_tables_use_visual_fallback(int rows, int columns) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Empty", rows, columns, 0d, emptyTile: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
    }

    [Theory]
    [InlineData((byte)0x78)]
    [InlineData((byte)0x7c)]
    [InlineData((byte)0xf8)]
    [InlineData((byte)0xfc)]
    public void Decimal128_special_encodings_disable_editable_reconstruction(byte highByte) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Special", 1, 1, 0d,
                decimal128SpecialHighByte: highByte)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(IWorkCellKind.Error, cell.Kind);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_CELL_DECODE");
    }

    private static byte[] CreatePngWithTrailingDeflateByte() {
        byte[] valid = ValidPreviewPng();
        const int idatChunkOffset = 33;
        int idatLength = valid[idatChunkOffset] << 24
            | valid[idatChunkOffset + 1] << 16
            | valid[idatChunkOffset + 2] << 8
            | valid[idatChunkOffset + 3];
        int dataOffset = idatChunkOffset + 8;
        byte[] imageData = valid[dataOffset..(dataOffset + idatLength)];
        byte[] malformedData = Message(imageData[..^4], new byte[] { 0 }, imageData[^4..]);
        int nextChunkOffset = dataOffset + idatLength + 4;
        return Message(valid[..idatChunkOffset], CreatePngChunk("IDAT", malformedData),
            valid[nextChunkOffset..]);
    }
}
