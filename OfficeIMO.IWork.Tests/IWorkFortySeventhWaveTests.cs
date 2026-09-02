using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_tile_fields_are_bounded_before_row_metadata_materialization() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Tile fields", 1, 1, 1d, unexpectedTileFieldCount: 7)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_TILE_FIELDS_UNSUPPORTED");
    }

    [Fact]
    public void Mixed_numbers_row_index_wire_kinds_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Mixed row", 1, 1, 1d, mixedRowIndexWire: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Repeated_pages_header_storage_disables_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithRepeatedHeaderStorage();

        IWorkPagesProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages).ReadPages();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_HEADER_FOOTER_DUPLICATE");
    }

    [Fact]
    public void Table_date_display_text_preserves_fractional_seconds() {
        DateTime date = new DateTime(2026, 8, 31, 14, 30, 15,
            DateTimeKind.Utc).AddTicks(1_234_567);
        var value = new IWorkTableCell(1, 1, IWorkCellKind.DateTime, date);
        var formula = new IWorkTableCell(1, 1, IWorkCellKind.Formula, date,
            formula: "=1", valueKind: IWorkCellKind.DateTime,
            formulaIsComplete: true);

        Assert.Equal("2026-08-31 14:30:15.1234567", value.DisplayText);
        Assert.Equal("2026-08-31 14:30:15.1234567", formula.CachedDisplayText);
    }

    private static MemoryStream CreatePagesPackageWithRepeatedHeaderStorage() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong sectionId = 3;
        const ulong headerFooterId = 4;
        const ulong headerStorageId = 5;
        byte[] sectionTable = Message(BytesField(1,
            Message(ReferenceField(2, sectionId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Body"), BytesField(17, sectionTable)),
                new[] { sectionId }),
            ArchiveRecord(sectionId, 10011,
                Message(ReferenceField(25, headerFooterId)),
                new[] { headerFooterId }),
            ArchiveRecord(headerFooterId, 10143,
                Message(ReferenceField(1, headerStorageId),
                    ReferenceField(1, headerStorageId)),
                new[] { headerStorageId }),
            ArchiveRecord(headerStorageId, 2001,
                Message(StringField(3, "Header"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
