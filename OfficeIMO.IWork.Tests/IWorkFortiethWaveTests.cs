using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(5)]
    [InlineData(6)]
    public void MessageInfo_semantic_fields_reject_mixed_invalid_wire_kinds(int field) {
        var messageInfoFields = new List<byte[]> {
            VarintField(1, 10000),
            VarintField(3, 0)
        };
        if (field is 2 or 5 or 6) messageInfoFields.Add(VarintField(field, 2));
        messageInfoFields.Add(FloatField(field, 1f));
        byte[] messageInfo = Message(messageInfoFields.ToArray());
        byte[] archiveInfo = Message(VarintField(1, 1), BytesField(2, messageInfo));
        byte[] stream = Message(Varint(checked((ulong)archiveInfo.Length)), archiveInfo);
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(stream)));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages));

        Assert.Contains("MessageInfo", exception.Message, StringComparison.Ordinal);
        Assert.Contains("wire encoding", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ArchiveInfo_identifiers_reject_mixed_invalid_wire_kinds() {
        byte[] messageInfo = Message(VarintField(1, 10000), VarintField(3, 0));
        byte[] archiveInfo = Message(
            VarintField(1, 1), FloatField(1, 1f), BytesField(2, messageInfo));
        byte[] stream = Message(Varint(checked((ulong)archiveInfo.Length)), archiveInfo);
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(stream)));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages));

        Assert.Contains("ArchiveInfo", exception.Message, StringComparison.Ordinal);
        Assert.Contains("identifier", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Undefined_conversion_modes_are_rejected_before_projection() {
        using MemoryStream package = CreatePagesPackage(
            includeBody: true, textBox: null, includePreview: false);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages);
        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            source.ToWordDocumentResult(
                new IWorkConversionOptions { Mode = (IWorkConversionMode)99 }));

        Assert.Equal("Mode", exception.ParamName);
    }

    [Fact]
    public void Undefined_document_kinds_are_rejected_before_package_reads() {
        using var package = new MemoryStream();

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            IWorkSourceDocument.Open(package, (IWorkDocumentKind)99));

        Assert.Equal("expectedKind", exception.ParamName);
    }

    [Fact]
    public void Undefined_projection_kinds_are_rejected() {
        using MemoryStream package = CreatePagesPackage(
            includeBody: true, textBox: null, includePreview: false);
        IWorkPagesProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages).ReadPages();

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            projection.CreateConversionReport((IWorkProjectionKind)99));

        Assert.Equal("projectionKind", exception.ParamName);
    }

    [Fact]
    public void Nonzero_decimal128_values_that_underflow_to_zero_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Underflow", 1, 1, 0d, decimal128Underflow: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(IWorkCellKind.Error, cell.Kind);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_DECODE");
    }
}
