using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Pages_body_storage_aliases_remain_editable_drawables() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong zOrderId = 3;
        const ulong shapeId = 4;
        byte[] drawable = Message(
            StringField(4, "https://example.test/body-alias"),
            StringField(8, "Aliased body text box"));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId), ReferenceField(20, zOrderId)),
                new[] { bodyId, zOrderId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Shared body"))),
            ArchiveRecord(zOrderId, 10020, Message(ReferenceField(1, shapeId)),
                new[] { shapeId }),
            ArchiveRecord(shapeId, 2011,
                Message(BytesField(1, Message(BytesField(1, drawable))),
                    ReferenceField(2, bodyId)),
                new[] { bodyId }));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);
        IWorkTextBox textBox = Assert.Single(result.Projection.TextBoxObjects);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
        Assert.Equal("Shared body", result.Projection.Body.PlainText);
        Assert.Equal("Shared body", textBox.Content.PlainText);
        Assert.Equal("https://example.test/body-alias", textBox.Hyperlink);
        Assert.Equal("Aliased body text box", textBox.AccessibilityDescription);
        Assert.Single(result.Document.Images);
    }

    [Fact]
    public void Pages_font_sizes_finer_than_half_a_point_use_visual_fallback() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(
            depth: 1, includePreview: true, fontSize: 10.25f);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
        Assert.Single(result.Document.Images);
    }

    [Fact]
    public void Mixed_wire_package_metadata_identifiers_disable_image_reconstruction() {
        using MemoryStream package = CreatePagesImagePackage(
            duplicateMetadata: false, imageCount: 1,
            imageBytes: ValidPreviewPng(), malformedIdentifierWire: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages);
        IWorkArchiveRecord image = Assert.Single(source.Records,
            record => record.MessageType == 3005);

        IWorkImageAsset? asset = IWorkDrawingReader.ReadImage(source, image,
            new IWorkProjectionBudget(new IWorkReadOptions()), out bool complete);

        Assert.Null(asset);
        Assert.False(complete);
    }

    [Fact]
    public void Mixed_wire_image_data_identifiers_disable_image_reconstruction() {
        using MemoryStream package = CreatePagesImagePackage(
            duplicateMetadata: false, imageCount: 1,
            imageBytes: ValidPreviewPng(), malformedImageIdentifierWire: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages);
        IWorkArchiveRecord image = Assert.Single(source.Records,
            record => record.MessageType == 3005);

        IWorkImageAsset? asset = IWorkDrawingReader.ReadImage(source, image,
            new IWorkProjectionBudget(new IWorkReadOptions()), out bool complete);

        Assert.Null(asset);
        Assert.False(complete);
    }

    [Fact]
    public void Keynote_font_sizes_finer_than_a_hundredth_point_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(
            1, fontSize: 10.125f);

        using var result = OfficeIMO.PowerPoint.PowerPointPresentation
            .LoadKeynoteWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
        Assert.Single(result.Document.Slides[0].Pictures);
    }

    [Theory]
    [InlineData(-1d)]
    [InlineData(0.5d)]
    [InlineData(2d)]
    public void Non_boolean_numbers_values_disable_editable_reconstruction(double value) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Malformed Boolean", 1, 1, value, boolean: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(IWorkCellKind.Error, cell.Kind);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_DECODE");
    }

    [Theory]
    [InlineData(0d, false)]
    [InlineData(1d, true)]
    public void Defined_numbers_boolean_values_remain_editable(double value, bool expected) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Boolean", 1, 1, value, boolean: true)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(IWorkCellKind.Boolean, cell.Kind);
        Assert.Equal(expected, cell.Value);
        Assert.Equal(expected,
            result.Document.Sheets[0].CellAt(1, 1).GetValue<bool>());
    }
}
