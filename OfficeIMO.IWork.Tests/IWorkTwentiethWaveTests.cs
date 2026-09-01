using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Rotated_pages_text_boxes_use_visual_fallback() {
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 36f), FloatField(2, 72f))),
            BytesField(2, Message(FloatField(1, 216f), FloatField(2, 108f))),
            FloatField(4, 15f));
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Rotated",
            includePreview: true, textBoxDrawable: Message(BytesField(1, geometry)));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Each_drawable_that_shares_text_storage_is_materialized(IWorkDocumentKind kind) {
        using MemoryStream package = CreatePackageWithSharedDrawableText(kind);

        if (kind == IWorkDocumentKind.Pages) {
            using var result = WordDocument.LoadPagesWithReport(package);
            Assert.False(result.IsVisualFallback);
            Assert.Equal(2, result.Projection.TextBoxObjects.Count);
            Assert.Equal(2, result.Document.TextBoxes.Count);
        } else {
            using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
            Assert.False(result.IsVisualFallback);
            Assert.Equal(2, Assert.Single(result.Projection.Slides).TextBoxes.Count);
            Assert.Equal(2, Assert.Single(result.Document.Slides).TextBoxes.Count());
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Shared_drawable_text_is_charged_for_each_destination_use(IWorkDocumentKind kind) {
        using MemoryStream package = CreatePackageWithSharedDrawableText(kind);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 11 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => {
            if (kind == IWorkDocumentKind.Pages) source.ReadPages();
            else source.ReadKeynote();
        });

        Assert.Contains("character count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Table_drawable_hyperlinks_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackageWithTableHyperlink();

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_HYPERLINK_UNSUPPORTED");
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Uncached_formulas_use_visual_fallback_in_noncalculating_owners(IWorkDocumentKind kind) {
        using MemoryStream package = CreatePackageWithUncachedFormulaTable(kind);

        if (kind == IWorkDocumentKind.Pages) {
            using var result = WordDocument.LoadPagesWithReport(package);
            IWorkTableCell cell = Assert.Single(Assert.Single(result.Projection.Tables).Cells);
            Assert.True(result.IsVisualFallback);
            Assert.True(result.Projection.HasEditableContent);
            Assert.True(cell.FormulaIsComplete);
            Assert.Null(cell.Value);
        } else {
            using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
            IWorkTableCell cell = Assert.Single(Assert.Single(
                Assert.Single(result.Projection.Slides).Tables).Cells);
            Assert.True(result.IsVisualFallback);
            Assert.True(result.Projection.HasEditableContent);
            Assert.True(cell.FormulaIsComplete);
            Assert.Null(cell.Value);
        }
    }

    private static MemoryStream CreatePackageWithSharedDrawableText(IWorkDocumentKind kind) {
        const ulong documentId = 1;
        const ulong bodyOrShowId = 2;
        const ulong nodeId = 3;
        const ulong slideOrFirstShapeId = 4;
        const ulong secondShapeId = 5;
        const ulong storageId = 6;
        byte[] firstDrawable = Message(BytesField(1, Message(
            BytesField(1, Message(FloatField(1, 10f), FloatField(2, 20f))),
            BytesField(2, Message(FloatField(1, 120f), FloatField(2, 40f))))));
        byte[] secondDrawable = Message(BytesField(1, Message(
            BytesField(1, Message(FloatField(1, 160f), FloatField(2, 20f))),
            BytesField(2, Message(FloatField(1, 120f), FloatField(2, 40f))))));
        byte[] firstShape = Message(BytesField(1, Message(BytesField(1, firstDrawable))),
            ReferenceField(2, storageId));
        byte[] secondShape = Message(BytesField(1, Message(BytesField(1, secondDrawable))),
            ReferenceField(2, storageId));

        if (kind == IWorkDocumentKind.Pages) {
            byte[] records = Message(
                ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyOrShowId)),
                    new[] { bodyOrShowId, slideOrFirstShapeId, secondShapeId }),
                ArchiveRecord(bodyOrShowId, 2001, Message(StringField(3, "Body"))),
                ArchiveRecord(slideOrFirstShapeId, 2011, firstShape, new[] { storageId }),
                ArchiveRecord(secondShapeId, 2011, secondShape, new[] { storageId }),
                ArchiveRecord(storageId, 2001, Message(StringField(3, "Shared"))));
            return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
        }

        const ulong slideId = 4;
        const ulong firstShapeId = 5;
        const ulong keynoteSecondShapeId = 7;
        byte[] keynoteRecords = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, bodyOrShowId))),
            ArchiveRecord(bodyOrShowId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5,
                Message(ReferenceField(7, firstShapeId), ReferenceField(7, keynoteSecondShapeId))),
            ArchiveRecord(firstShapeId, 2011, firstShape, new[] { storageId }),
            ArchiveRecord(keynoteSecondShapeId, 2011, secondShape, new[] { storageId }),
            ArchiveRecord(storageId, 2001, Message(StringField(3, "Shared"))));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(keynoteRecords)));
    }

    private static MemoryStream CreateNumbersPackageWithTableHyperlink() {
        const ulong documentId = 1;
        const ulong sheetId = 2;
        const ulong tableId = 10;
        const ulong modelId = 11;
        byte[] drawable = Message(StringField(4, "https://example.test/table"));
        byte[] model = Message(
            BytesField(4, Message(BytesField(3, Message()))),
            VarintField(6, 1), VarintField(7, 1), StringField(8, "Linked"));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(1, sheetId))),
            ArchiveRecord(sheetId, 2,
                Message(StringField(1, "Sheet"), ReferenceField(2, tableId))),
            ArchiveRecord(tableId, 6000,
                Message(BytesField(1, drawable), ReferenceField(2, modelId)), new[] { modelId }),
            ArchiveRecord(modelId, 6001, model));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreatePackageWithUncachedFormulaTable(IWorkDocumentKind kind) {
        const ulong documentId = 1;
        const ulong bodyOrShowId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong tableId = 10;
        const ulong modelId = 11;
        const ulong tileId = 12;
        const ulong formulaListId = 13;
        var table = new TableSpec("Formula", 1, 1, 0d,
            hasFormula: true, formulaWithoutCachedValue: true);
        byte[] row = CreateBncRow(table);
        byte[] tileEntry = Message(VarintField(1, 0), ReferenceField(2, tileId));
        byte[] store = Message(
            BytesField(3, Message(BytesField(1, tileEntry))),
            ReferenceField(6, formulaListId));
        byte[] model = Message(BytesField(4, store),
            VarintField(6, 1), VarintField(7, 1), StringField(8, "Formula"));
        byte[] formulaEntry = Message(VarintField(1, 0), BytesField(5, FormulaConstant(1d)));
        var records = new List<byte[]>();
        if (kind == IWorkDocumentKind.Pages) {
            records.Add(ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyOrShowId)), new[] { bodyOrShowId, tableId }));
            records.Add(ArchiveRecord(bodyOrShowId, 2001, Message(StringField(3, "Body"))));
        } else {
            records.Add(ArchiveRecord(documentId, 1, Message(ReferenceField(2, bodyOrShowId))));
            records.Add(ArchiveRecord(bodyOrShowId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))));
            records.Add(ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))));
            records.Add(ArchiveRecord(slideId, 5, Message(ReferenceField(6, tableId))));
        }
        records.Add(ArchiveRecord(tableId, 6000,
            kind == IWorkDocumentKind.Keynote
                ? Message(BytesField(1, GeometryDrawable(72f, 72f, 120f, 40f)),
                    ReferenceField(2, modelId))
                : Message(ReferenceField(2, modelId)),
            new[] { modelId }));
        records.Add(ArchiveRecord(modelId, 6001, model, new[] { tileId, formulaListId }));
        records.Add(ArchiveRecord(tileId, 6002, Message(BytesField(5, row))));
        records.Add(ArchiveRecord(formulaListId, 6201, Message(BytesField(3, formulaEntry))));
        return CreatePackage(
            (kind == IWorkDocumentKind.Pages ? "Index/Document.iwa" : "Index/Slide.iwa",
                FrameIwa(Message(records.ToArray()))),
            ("preview.png", ValidPreviewPng()));
    }
}
