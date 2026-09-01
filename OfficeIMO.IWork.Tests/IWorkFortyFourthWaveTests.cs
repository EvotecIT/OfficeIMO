using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Mixed_formula_node_type_wires_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula", 1, 1, 1d, hasFormula: true,
                completeFormula: true, mixedFormulaTypeWire: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.False(cell.FormulaIsComplete);
        Assert.False(result.IsVisualFallback);
        Assert.Null(result.Document.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(1d, result.Document.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Numbers_row_heights_preserve_source_precision_in_the_excel_owner() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Height", 1, 1, 1d, defaultRowHeight: 12.345d)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(12.345d, result.Document.Sheets[0].DefaultRowHeight);
    }

    [Fact]
    public void Pages_layout_measurements_finer_than_one_twip_use_visual_fallback() {
        byte[] layout = Message(
            FloatField(30, 612.025f), FloatField(31, 792f),
            FloatField(32, 72f), FloatField(33, 72f),
            FloatField(34, 72f), FloatField(35, 72f),
            FloatField(36, 36f), FloatField(37, 36f));
        using MemoryStream package = CreatePagesPackage(
            includeBody: true, textBox: null, includePreview: true,
            documentLayoutFields: layout);

        using var result = WordIWorkConverter.LoadPagesWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }

    [Fact]
    public void Slide_named_archives_do_not_override_a_pages_root() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Pages body"))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("Index/Slide-decoy.iwa", FrameIwa(ArchiveRecord(900, 999, Message()))));

        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages);

        Assert.Equal(IWorkDocumentKind.Pages, source.Kind);
        Assert.Equal("Pages body", source.ReadPages().Body.PlainText);
    }

    [Fact]
    public void Pages_z_order_references_are_bounded_before_nested_materialization() {
        using MemoryStream package = CreatePagesDrawableOccurrencePackage(
            duplicateWithinField: false, floating: false, occurrenceCount: 13);
        var options = new IWorkReadOptions {
            MaximumProjectedImages = 1,
            MaximumProjectedTables = 1,
            MaximumProjectedTextItems = 10
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            WordIWorkConverter.LoadPagesWithReport(package, options));

        Assert.Contains("drawable references", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Numbers_drawable_references_share_the_projection_bound() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Repeated", 1, 1, 1d)
        }, repeatedFirstDrawableCount: 3);
        var options = new IWorkReadOptions {
            MaximumProjectedImages = 1,
            MaximumProjectedTables = 1,
            MaximumProjectedTextItems = 1
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            ExcelIWorkConverter.LoadNumbersWithReport(package, options));

        Assert.Contains("drawable references", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Keynote_drawable_references_share_the_projection_bound() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(
            1, drawableReferenceCount: 13);
        var options = new IWorkReadOptions {
            MaximumProjectedImages = 1,
            MaximumProjectedTables = 1,
            MaximumProjectedTextItems = 10
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            PowerPointIWorkConverter.LoadKeynoteWithReport(package, options));

        Assert.Contains("drawable references", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }
}
