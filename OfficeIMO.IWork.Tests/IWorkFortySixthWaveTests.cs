using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Formula_node_array_total_fields_are_bounded_before_materialization() {
        byte[] node = Message(VarintField(1, 17), DoubleField(4, 1d));
        byte[] formula = Message(BytesField(1, Message(
            BytesField(1, node), VarintField(2, 1))));
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula", 1, 1, 1d, hasFormula: true,
                formulaPayload: formula)
        });

        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumFormulaNodes = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => source.ReadNumbers());
        Assert.Contains("syntax-node limit", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Image_metadata_entries_are_bounded_before_materialization() {
        using MemoryStream package = CreatePagesImagePackage(
            duplicateMetadata: false, imageCount: 1,
            imageBytes: ValidPreviewPng(), metadataEntryCount: 2);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumImageMetadataEntries = 1 });
        IWorkArchiveRecord image = Assert.Single(source.Records,
            record => record.MessageType == 3005);
        var budget = new IWorkProjectionBudget(source.Options);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => IWorkDrawingReader.ReadImage(source, image, budget, out _));
        Assert.Contains("image metadata", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Numbers_sheet_references_are_bounded_during_kind_detection() {
        using MemoryStream package = CreateNumbersPackage(
            Array.Empty<TableSpec>(), sheetReferenceCount: 2);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumProjectedSheets = 1 }));
        Assert.Contains("sheet count", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

}
