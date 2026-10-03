using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Text_only_owners_reject_formula_caches_with_incomplete_styles(
        IWorkDocumentKind kind) {
        using MemoryStream package = CreateFormulaTableWithRichCacheStyle(kind);

        if (kind == IWorkDocumentKind.Pages) {
            using var result = WordIWorkConverter.ConvertPagesToWordResult(package);
            IWorkTableCell cell = Assert.Single(Assert.Single(result.Projection.Tables).Cells);
            Assert.True(cell.FormulaIsComplete);
            Assert.True(cell.CachedValueIsComplete);
            Assert.False(cell.RichText!.IsComplete);
            Assert.True(result.Projection.HasEditableContent);
            Assert.True(result.IsVisualFallback);
            AssertSourceFormula(result.Report);
        } else {
            using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
            IWorkTableCell cell = Assert.Single(Assert.Single(
                Assert.Single(result.Projection.Slides).Tables).Cells);
            Assert.True(cell.FormulaIsComplete);
            Assert.True(cell.CachedValueIsComplete);
            Assert.False(cell.RichText!.IsComplete);
            Assert.True(result.Projection.HasEditableContent);
            Assert.True(result.IsVisualFallback);
            AssertSourceFormula(result.Report);
        }
        static void AssertSourceFormula(IWorkConversionReport report) {
            IWorkFormulaCellStatus assessment = Assert.Single(report.FormulaCells);
            Assert.True(assessment.ExpressionIsComplete);
            Assert.Equal(IWorkFormulaCacheStatus.Complete, assessment.CacheStatus);
            Assert.Equal(10ul, assessment.TableIdentity!.RecordIdentifier);
        }
    }

    private static MemoryStream CreateFormulaTableWithRichCacheStyle(
        IWorkDocumentKind kind, int? paginationStyleField = null) {
        var cell = new byte[20];
        cell[0] = 5;
        cell[1] = 9;
        WriteUInt32(cell, 8, (1u << 4) | (1u << 9));
        WriteUInt32(cell, 12, 1);
        byte[] row = Message(VarintField(1, 0), BytesField(6, cell),
            BytesField(7, new byte[] { 0, 0 }));
        byte[] store = Message(BytesField(3, Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, 12))))),
            ReferenceField(6, 13), ReferenceField(17, 14));
        byte[] records = Message(
            kind == IWorkDocumentKind.Pages
                ? ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)),
                    new ulong[] { 2, 10 })
                : ArchiveRecord(1, 1, Message(ReferenceField(2, 2)), new ulong[] { 2 }),
            kind == IWorkDocumentKind.Pages
                ? ArchiveRecord(2, 2001, Message(StringField(3, "Body")))
                : ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3))),
                    new ulong[] { 3 }),
            kind == IWorkDocumentKind.Pages ? Array.Empty<byte>()
                : ArchiveRecord(3, 4, Message(ReferenceField(2, 4)), new ulong[] { 4 }),
            kind == IWorkDocumentKind.Pages ? Array.Empty<byte>()
                : ArchiveRecord(4, 5, Message(ReferenceField(6, 10)), new ulong[] { 10 }),
            ArchiveRecord(10, 6000,
                kind == IWorkDocumentKind.Pages
                    ? Message(ReferenceField(2, 11))
                    : Message(BytesField(1, GeometryDrawable(72f, 72f, 120f, 40f)),
                        ReferenceField(2, 11)), new ulong[] { 11 }),
            ArchiveRecord(11, 6001, Message(BytesField(4, store), VarintField(6, 1),
                VarintField(7, 1), StringField(8, "Styled formula")),
                new ulong[] { 12, 13, 14 }),
            ArchiveRecord(12, 6002, Message(BytesField(5, row))),
            ArchiveRecord(13, 6201, Message(VarintField(1, 3), BytesField(3,
                Message(VarintField(1, 0), BytesField(5, FormulaConstant(1d)))))),
            ArchiveRecord(14, 6005, Message(VarintField(1, 8), BytesField(3,
                Message(VarintField(1, 1), ReferenceField(9, 15)))), new ulong[] { 15 }),
            ArchiveRecord(15, 6218, Message(ReferenceField(1, 16)), new ulong[] { 16 }),
            ArchiveRecord(16, 2001, Message(StringField(3, "Styled"),
                BytesField(5, Message(BytesField(1, Message(VarintField(1, 0),
                    ReferenceField(2, paginationStyleField.HasValue ? 17UL : 99UL))))))),
            paginationStyleField.HasValue
                ? ArchiveRecord(17, 2022, Message(BytesField(12, Message(VarintField(paginationStyleField.Value, 1)))))
                : Array.Empty<byte>());
        return CreatePackage(
            (kind == IWorkDocumentKind.Pages ? "Index/Document.iwa" : "Index/Slide.iwa",
                FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
