using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Archive_references_are_bounded_across_all_records_and_reference_kinds() {
        byte[] firstMessageInfo = Message(
            VarintField(1, 1),
            BytesField(2, new byte[] { 1 }),
            VarintField(3, 0),
            BytesField(5, new byte[] { 2 }));
        byte[] secondMessageInfo = Message(
            VarintField(1, 2),
            VarintField(3, 0),
            BytesField(6, new byte[] { 3, 4 }));
        byte[] firstArchiveInfo = Message(VarintField(1, 1),
            BytesField(2, firstMessageInfo));
        byte[] secondArchiveInfo = Message(VarintField(1, 2),
            BytesField(2, secondMessageInfo));
        byte[] iwa = Message(
            Varint(checked((ulong)firstArchiveInfo.Length)), firstArchiveInfo,
            Varint(checked((ulong)secondArchiveInfo.Length)), secondArchiveInfo);
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(iwa)));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumArchiveReferenceCount = 3 }));

        Assert.Contains("remaining aggregate limit", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Repeated_hyperlink_targets_are_charged_for_each_projected_run() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong hyperlinkId = 3;
        const string target = "https://example.test/repeated";
        byte[] hyperlinkTable = Message(
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, hyperlinkId))),
            BytesField(1, Message(VarintField(1, 1), ReferenceField(2, hyperlinkId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)),
                new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "AB"), BytesField(11, hyperlinkTable)),
                new[] { hyperlinkId }),
            ArchiveRecord(hyperlinkId, 2032, Message(StringField(2, target))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages,
            new IWorkReadOptions {
                MaximumProjectedTextCharacters = 2 + target.Length * 2 - 1
            });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ReadPages());

        Assert.Contains("Text character count", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Unresolved_formula_text_caches_disable_editable_numbers_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula cache", 1, 1, 0d, hasFormula: true,
                textValue: "Cached", completeFormula: true, missingStringEntry: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(
            result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(IWorkCellKind.Formula, cell.Kind);
        Assert.Null(cell.Value);
        Assert.False(cell.FormulaIsComplete);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_FORMULA_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_table_geometry_preserves_authored_extent_when_defaults_differ() {
        using MemoryStream package = CreateKeynotePackageWithTableDefaults(
            rows: 2, columns: 2, defaultRowHeight: 5d, defaultColumnWidth: 10d,
            tableDrawable: GeometryDrawable(10f, 20f, 60f, 40f));

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        PowerPointTable table = Assert.Single(Assert.Single(result.Document.Slides).Tables);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(10d, table.LeftPoints, 5);
        Assert.Equal(20d, table.TopPoints, 5);
        Assert.Equal(60d, table.WidthPoints, 5);
        Assert.Equal(40d, table.HeightPoints, 5);
        Assert.Equal(30d, table.GetColumnWidthPoints(0), 5);
        Assert.Equal(20d, table.GetRowHeightPoints(0), 5);
    }

    [Fact]
    public void Jpeg_assets_reject_payload_after_the_first_end_marker() {
        using FileStream input = File.OpenRead(Fixture("nim-iwork/simple.pages"));
        using var fixture = new System.IO.Compression.ZipArchive(input,
            System.IO.Compression.ZipArchiveMode.Read, leaveOpen: false);
        byte[] jpeg = ReadEntry(fixture, "preview.jpg");
        byte[] trailed = jpeg.Concat(new byte[] { 1, 2, 3, 0xff, 0xd9 }).ToArray();

        (int? width, int? height) = IWorkImageInfo.Read(
            trailed, "image/jpeg", 64L * 1024 * 1024);

        Assert.Null(width);
        Assert.Null(height);
    }

    [Fact]
    public void Malformed_pages_roots_use_visual_fallback() {
        using MemoryStream package = CreateMalformedRootPackage(10000);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DOCUMENT_MALFORMED");
    }

    [Fact]
    public void Malformed_numbers_roots_use_visual_fallback() {
        using MemoryStream package = CreateMalformedRootPackage(1);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_DOCUMENT_MALFORMED");
    }

    [Fact]
    public void Malformed_keynote_roots_use_visual_fallback() {
        using MemoryStream package = CreateMalformedRootPackage(1);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_DOCUMENT_MALFORMED");
    }

    private static MemoryStream CreateMalformedRootPackage(uint rootType) {
        byte[] records = ArchiveRecord(1, rootType, new byte[] { 0x80 });
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
