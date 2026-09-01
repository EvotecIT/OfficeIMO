using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Metadata_only_pages_text_boxes_respect_the_text_item_budget() {
        using MemoryStream package = CreatePagesPackageWithMetadataOnlyTextBoxes();
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextItems = 3 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadPages());

        Assert.Contains("text item count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Metadata_only_keynote_text_boxes_respect_the_text_item_budget() {
        using MemoryStream package = CreateKeynotePackageWithMetadataOnlyTextBoxes();
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote,
            new IWorkReadOptions { MaximumProjectedTextItems = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadKeynote());

        Assert.Contains("text item count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("(3)", 3, "decimal")]
    [InlineData("(c)", 3, "lowerLetter")]
    [InlineData("(iv)", 4, "lowerRoman")]
    public void Parenthesized_pages_lists_preserve_native_numbering(
        string label, int expectedStart, string expectedFormat) {
        using MemoryStream package = CreatePagesPackageWithListLabel(label);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        Numbering numbering = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering
            ?? throw new InvalidDataException("The reconstructed DOCX has no numbering definitions.");
        Level level = Assert.Single(numbering.Elements<AbstractNum>().SelectMany(item =>
            item.Elements<Level>()));

        Assert.Equal(expectedStart, level.StartNumberingValue?.Val?.Value);
        NumberFormatValues format = expectedFormat switch {
            "decimal" => NumberFormatValues.Decimal,
            "lowerLetter" => NumberFormatValues.LowerLetter,
            _ => NumberFormatValues.LowerRoman
        };
        Assert.Equal(format, level.NumberingFormat?.Val?.Value);
        Assert.Equal("(%1)", level.LevelText?.Val?.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Populated_numbers_offsets_beyond_declared_columns_disable_editable_reconstruction(
        bool wideOffsets) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Out of range", 1, 1, 42d, wideOffsets: wideOffsets,
                populatedOffsetBeyondColumns: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Empty_numbers_offsets_beyond_declared_columns_remain_editable() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Trailing empty", 1, 1, 42d, emptyOffsetBeyondColumns: true)
        });

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(42d, result.Value.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(4, "https://example.test/")]
    [InlineData(8, "Accessible shape")]
    public void Numbers_text_shape_metadata_disables_lossy_editable_reconstruction(
        int field, string value) {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(), textBox: "Linked",
            includePreview: true, textBoxDrawable: Message(StringField(field, value)));

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_NUMBERS_TEXT_METADATA_UNSUPPORTED");
    }

    [Fact]
    public void Malformed_numbers_text_shape_drawables_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(), textBox: "Text",
            includePreview: true, textBoxDrawable: new byte[] { 0x08, 0x80 });

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_NUMBERS_TEXT_METADATA_UNSUPPORTED");
    }

    private static MemoryStream CreatePagesPackageWithMetadataOnlyTextBoxes() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstShapeId = 3;
        const ulong secondShapeId = 4;
        const ulong emptyStorageId = 5;
        byte[] firstShape = Message(
            BytesField(1, Message(BytesField(1, Message(StringField(8, "First"))))),
            ReferenceField(2, emptyStorageId));
        byte[] secondShape = Message(
            BytesField(1, Message(BytesField(1,
                Message(GeometryDrawable(10f, 10f, 100f, 50f),
                    StringField(8, "Second"))))),
            ReferenceField(2, emptyStorageId));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)),
                new[] { bodyId, firstShapeId, secondShapeId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(firstShapeId, 2011, firstShape, new[] { emptyStorageId }),
            ArchiveRecord(secondShapeId, 2011, secondShape, new[] { emptyStorageId }),
            ArchiveRecord(emptyStorageId, 2001, Message(StringField(3, string.Empty))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreateKeynotePackageWithMetadataOnlyTextBoxes() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong firstShapeId = 5;
        const ulong secondShapeId = 6;
        const ulong emptyStorageId = 7;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        byte[] firstShape = Message(
            BytesField(1, Message(BytesField(1, Message(StringField(8, "First"))))),
            ReferenceField(2, emptyStorageId));
        byte[] secondShape = Message(
            BytesField(1, Message(BytesField(1,
                Message(GeometryDrawable(10f, 10f, 100f, 50f),
                    StringField(8, "Second"))))),
            ReferenceField(2, emptyStorageId));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2, KeynoteShow(slideTree)),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5,
                Message(ReferenceField(5, firstShapeId), ReferenceField(7, secondShapeId))),
            ArchiveRecord(firstShapeId, 2011, firstShape, new[] { emptyStorageId }),
            ArchiveRecord(secondShapeId, 2011, secondShape, new[] { emptyStorageId }),
            ArchiveRecord(emptyStorageId, 2001, Message(StringField(3, string.Empty))));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
