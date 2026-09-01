using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Keynote_show_rejects_repeated_slide_trees() {
        byte[] firstTree = Message(ReferenceField(2, 3));
        byte[] secondTree = Message(ReferenceField(2, 5));
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2,
                Message(BytesField(3, firstTree), BytesField(3, secondTree))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message()),
            ArchiveRecord(5, 4, Message(ReferenceField(2, 6))),
            ArchiveRecord(6, 5, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SLIDE_TREE_MISSING");
    }

    [Fact]
    public void Numbers_tables_reject_repeated_cell_stores() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Repeated store", 1, 1, 42d, duplicateTableStore: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Pages_document_rejects_repeated_z_order_references() {
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2), ReferenceField(20, 3), ReferenceField(20, 4))),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10021, Message(ReferenceField(1, 5))),
            ArchiveRecord(4, 10021, Message(ReferenceField(1, 7))),
            ArchiveRecord(5, 2011, Message(ReferenceField(2, 6))),
            ArchiveRecord(6, 2001, Message(StringField(3, "First"))),
            ArchiveRecord(7, 2011, Message(ReferenceField(2, 8))),
            ArchiveRecord(8, 2001, Message(StringField(3, "Second"))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Formula_reader_rejects_repeated_node_arrays() {
        byte[] first = Message(BytesField(1,
            Message(VarintField(1, 17), DoubleField(4, 1d))));
        byte[] second = Message(BytesField(1,
            Message(VarintField(1, 17), DoubleField(4, 2d))));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, first), BytesField(1, second)),
            new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(
            formula, 0, 0, maximumNodes: 10, maximumCharacters: 100);

        Assert.False(result.IsComplete);
        Assert.Equal(string.Empty, result.Text);
    }

    [Fact]
    public void Numbered_list_levels_require_a_corresponding_label() {
        byte[] listTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, 3))));
        byte[] records = Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2))),
            ArchiveRecord(2, 2001,
                Message(StringField(3, "Item"), BytesField(7, listTable))),
            ArchiveRecord(3, 2023, Message(VarintField(11, 1))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
    }
}
