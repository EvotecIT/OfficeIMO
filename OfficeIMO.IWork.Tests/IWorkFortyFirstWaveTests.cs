using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Mixed_text_storage_wire_kinds_disable_editable_reconstruction() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Body"), VarintField(3, 1))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Mixed_page_layout_wire_kinds_disable_editable_reconstruction() {
        byte[] layout = Message(
            PageLayoutFields(72f),
            VarintField(30, 612));
        using MemoryStream package = CreatePagesPackage(
            includeBody: true, textBox: null, includePreview: true,
            documentLayoutFields: layout);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_LAYOUT_UNSUPPORTED");
    }

    [Fact]
    public void Unknown_numbers_cell_value_flags_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Unknown cell value", 1, 1, 42d,
                unknownCellValueFlag: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(IWorkCellKind.Error, cell.Kind);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_DECODE");
    }

    [Fact]
    public void Duplicate_pages_z_order_occurrences_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesDrawableOccurrencePackage(
            duplicateWithinField: true, floating: false);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Duplicate_pages_floating_field_occurrences_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesDrawableOccurrencePackage(
            duplicateWithinField: true, floating: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Pages_drawable_aliases_across_discovery_paths_remain_editable() {
        using MemoryStream package = CreatePagesDrawableOccurrencePackage(
            duplicateWithinField: false, floating: true, aliasAcrossFloatingFields: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("Shape", Assert.Single(result.Projection.TextBoxes));
    }

    private static MemoryStream CreatePagesDrawableOccurrencePackage(bool duplicateWithinField,
        bool floating, bool aliasAcrossFloatingFields = false, int? occurrenceCount = null) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong orderId = 3;
        const ulong shapeId = 4;
        const ulong storageId = 5;
        byte[] occurrence = ReferenceField(1, shapeId);
        byte[] orderedPayload;
        byte[] documentOrderReference;
        if (floating) {
            byte[] entry = Message(occurrence);
            byte[] pageGroup = Message(
                BytesField(2, entry),
                duplicateWithinField ? BytesField(2, entry) : Array.Empty<byte>(),
                aliasAcrossFloatingFields ? BytesField(3, entry) : Array.Empty<byte>());
            orderedPayload = Message(BytesField(1, pageGroup));
            documentOrderReference = ReferenceField(3, orderId);
        } else {
            orderedPayload = occurrenceCount.HasValue
                ? Message(Enumerable.Range(0, occurrenceCount.Value).Select(_ => occurrence).ToArray())
                : Message(occurrence,
                    duplicateWithinField ? occurrence : Array.Empty<byte>());
            documentOrderReference = ReferenceField(20, orderId);
        }
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId), documentOrderReference),
                new[] { bodyId, orderId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(orderId, 10020, orderedPayload, new[] { shapeId }),
            ArchiveRecord(shapeId, 2011,
                Message(BytesField(1, Message(BytesField(1,
                            GeometryDrawable(10f, 10f, 100f, 50f)))),
                    ReferenceField(2, storageId)), new[] { storageId }),
            ArchiveRecord(storageId, 2001, Message(StringField(3, "Shape"))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
