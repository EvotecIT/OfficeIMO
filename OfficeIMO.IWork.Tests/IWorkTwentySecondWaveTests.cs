using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(5)]
    [InlineData(7)]
    [InlineData(8)]
    [InlineData(11)]
    public void Rich_text_boundaries_beyond_the_text_use_visual_fallback(int tableField) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong referencedId = 3;
        byte[] attributeTable = Message(BytesField(1,
            Message(VarintField(1, 3), ReferenceField(2, referencedId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "AB"), BytesField(tableField, attributeTable)),
                new[] { referencedId }),
            ArchiveRecord(referencedId, 2021, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Body.IsComplete);
    }

    [Theory]
    [InlineData(16, -1d)]
    [InlineData(17, double.NaN)]
    public void Invalid_declared_table_default_sizes_use_visual_fallback(int field, double value) {
        using MemoryStream package = CreateNumbersPackageWithDeclaredTableSize(field, value);

        using var result = OfficeIMO.Excel.ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_DIMENSIONS_UNSUPPORTED");
    }

    [Fact]
    public void Pages_table_default_dimensions_are_applied_to_word() {
        using MemoryStream package = CreatePagesPackageWithTableGeometry(
            0f, 0f, 0f, 0f, 0f, includePreview: false,
            defaultRowHeight: 20d, defaultColumnWidth: 40d);

        using var result = WordDocument.LoadPagesWithReport(package);
        WordTable table = Assert.Single(result.Document.Tables);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(400, Assert.Single(table.RowHeight));
        Assert.Equal(800, Assert.Single(table.ColumnWidth));
    }

    [Fact]
    public void Small_keynote_table_defaults_are_preserved_without_minimum_inflation() {
        using MemoryStream package = CreateKeynotePackageWithTableDefaults(
            rows: 2, columns: 2, defaultRowHeight: 10d, defaultColumnWidth: 30d);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        PowerPointTable table = Assert.Single(Assert.Single(result.Document.Slides).Tables);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(60d, table.WidthPoints, 5);
        Assert.Equal(20d, table.HeightPoints, 5);
        Assert.Equal(30d, table.GetColumnWidthPoints(0), 5);
        Assert.Equal(10d, table.GetRowHeightPoints(0), 5);
    }

    [Fact]
    public void Keynote_list_levels_above_eight_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithListLevel(9);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(9, Assert.Single(Assert.Single(
            result.Projection.Slides).TitleBox!.Content.Paragraphs).ListLevel);
    }

    private static MemoryStream CreateNumbersPackageWithDeclaredTableSize(int field, double value) {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        const ulong tableId = 10;
        const ulong modelId = 11;
        byte[] model = Message(
            BytesField(4, Message(BytesField(3, Message()))),
            VarintField(6, 1), VarintField(7, 1), StringField(8, "Invalid size"),
            DoubleField(field, value));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(1, sheetId))),
            ArchiveRecord(sheetId, 2,
                Message(StringField(1, "Sheet"), ReferenceField(2, tableId))),
            ArchiveRecord(tableId, 6000,
                Message(ReferenceField(2, modelId)), new[] { modelId }),
            ArchiveRecord(modelId, 6001, model));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreateKeynotePackageWithTableDefaults(
        int rows, int columns, double defaultRowHeight, double defaultColumnWidth,
        bool includePreview = false, string? accessibilityDescription = null,
        byte[]? tableDrawable = null) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong tableId = 10;
        const ulong modelId = 11;
        byte[] model = Message(
            BytesField(4, Message(BytesField(3, Message()))),
            VarintField(6, checked((ulong)rows)), VarintField(7, checked((ulong)columns)),
            StringField(8, "Small"), DoubleField(16, defaultRowHeight),
            DoubleField(17, defaultColumnWidth));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(6, tableId))),
            ArchiveRecord(tableId, 6000,
                Message(
                    tableDrawable == null
                        ? Array.Empty<byte>()
                        : BytesField(1, tableDrawable),
                    accessibilityDescription == null
                        ? Array.Empty<byte>()
                        : BytesField(1, Message(StringField(8, accessibilityDescription))),
                    ReferenceField(2, modelId)), new[] { modelId }),
            ArchiveRecord(modelId, 6001, model));
        return includePreview
            ? CreatePackage(("Index/Slide.iwa", FrameIwa(records)),
                ("preview.png", ValidPreviewPng()))
            : CreatePackage(("Index/Slide.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreateKeynotePackageWithListLevel(int level) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        const ulong listStyleId = 7;
        const ulong paragraphStyleId = 8;
        byte[] listTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, listStyleId))));
        byte[] paragraphTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, paragraphStyleId))));
        var listFields = new List<byte[]>();
        for (int index = 0; index <= level; index++) {
            listFields.Add(VarintField(11, 1));
            listFields.Add(FloatField(13, index * 18f));
        }
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "Deep"), BytesField(5, paragraphTable),
                    BytesField(7, listTable)), new[] { listStyleId, paragraphStyleId }),
            ArchiveRecord(listStyleId, 2023, Message(listFields.ToArray())),
            ArchiveRecord(paragraphStyleId, 2022,
                Message(BytesField(12, Message(FloatField(11, level * 18f))))));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
