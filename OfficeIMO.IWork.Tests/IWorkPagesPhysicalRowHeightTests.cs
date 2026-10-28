using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_Pages_physical_row_heights_exclude_Word_added_bottom_padding() {
        using var result = WordIWorkConverter.ConvertPagesToWordResult(
            CorpusFixture("native-exports/pages-table-borders-v15.4.pages"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        IWorkTable source = Assert.Single(result.Projection.Tables);
        Assert.Equal(new double?[] { 32, 31, 32, 32 }, Enumerable.Range(1, 4).Select(source.GetRowHeight));
        result.Report.RequireCompleteEditableReconstruction();

        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        WordTable table = Assert.Single(reopened.Tables);
        Assert.Equal(new[] { 560, 540, 560, 560 }, table.RowHeight);
        for (int row = 0; row < table.Rows.Count; row++) {
            var cells = table.Rows[row].Cells;
            Assert.All(cells, cell => Assert.Equal((short)80, checked((short?)(cell.MarginBottom?.ToTwips()))));
            Assert.Equal(source.GetRowHeight(row + 1), (checked((int?)(table.Rows[row].Height?.ToTwips()))!.Value + 80) / 20d);
        }
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData("ordinary", 240, 1)]
    [InlineData("horizontal", 360, 1)]
    [InlineData("vertical", 240, 2)]
    public void Pages_exact_row_height_uses_largest_active_bottom_padding_and_ignores_top(string merge, int expected, int rowCount) {
        // The second cell is genuinely blank. Its selected padding still affects an ordinary row,
        // but a covered horizontal cell has no independent Word layout box.
        using var package = PhysicalRowPackage(20, 2, 8, merge, rowCount);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        Assert.Equal(IWorkCellKind.Empty, result.Projection.Tables[0].GetCell(1, 2)!.Kind);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        var table = Assert.Single(reopened.Tables);
        Assert.Equal(OfficeIMO.OfficeLength.Points(expected / 20D), table.Rows[0].Height);
        if (rowCount == 2) Assert.Equal(OfficeIMO.OfficeLength.Points(20), table.Rows[1].Height);
        Assert.Equal((short)140, checked((short?)(table.Rows[0].Cells[0].MarginTop?.ToTwips())));
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void Pages_partial_row_height_subtracts_serialized_rounded_padding() {
        using var package = PhysicalRowPackage(.125, .0625f, .0625f);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_PAGES_DOCX_PRECISION");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        var row = Assert.Single(reopened.Tables).Rows[0];
        Assert.Equal(OfficeIMO.OfficeLength.Points(0.1), row.Height); // Physical 3 twips, less the serialized 1-twip bottom padding.
        Assert.All(row.Cells, cell => Assert.Equal((short)1, checked((short?)(cell.MarginBottom?.ToTwips()))));
    }

    [Fact]
    public void Pages_exact_rows_retain_vertical_merge_continuation_bottom_padding() {
        using var package = PhysicalRowPackage(20, 2, 8, "vertical-styled", 2);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        Assert.Equal(8, result.Projection.Tables[0].GetCell(2, 1)!.Padding!.BottomPoints);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Assert.Equal(new[] { 240, 240 }, Assert.Single(reopened.Tables).RowHeight);
        Assert.Equal(WordCellMerge.Continue, reopened.Tables[0].Rows[1].Cells[0].VerticalMerge);
        Assert.Equal((short)160, checked((short?)(reopened.Tables[0].Rows[1].Cells[0].MarginBottom?.ToTwips())));
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void Pages_physical_row_box_can_reach_Word_declaration_limit_plus_padding() {
        using var package = PhysicalRowPackage(1588, 4, 4);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Assert.Equal(OfficeIMO.OfficeLength.Points(1584), Assert.Single(reopened.Tables).Rows[0].Height);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(4, 4, false)]
    [InlineData(4, 4, true)]
    [InlineData(2000, 4, false)]
    [InlineData(2000, 4, true)]
    public void Pages_unrepresentable_exact_row_box_is_rejected_before_DOCX_creation(double height, float padding, bool partial) {
        using var package = PhysicalRowPackage(height, padding, padding);
        Assert.Contains("row height", Assert.Throws<InvalidDataException>(() =>
            WordIWorkConverter.ConvertPagesToWordResult(package,
                conversionOptions: new IWorkConversionOptions {
                    Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = partial
                })).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Pages_minimum_row_constraint_also_respects_Word_declaration_range(bool partial) {
        using var package = PhysicalRowPackage(2000, 4, 4, automatic: true);
        Assert.Contains("row height", Assert.Throws<InvalidDataException>(() =>
            WordIWorkConverter.ConvertPagesToWordResult(package,
                conversionOptions: new IWorkConversionOptions {
                    Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = partial
                })).Message, StringComparison.OrdinalIgnoreCase);
    }

    private static MemoryStream PhysicalRowPackage(double height, float firstBottom, float secondBottom,
        string merge = "ordinary", int rowCount = 1, bool automatic = false) =>
        CellFillPackage(new[] {
            CellLayoutStyle(30, BytesField(9, Message(FloatField(2, 7), FloatField(4, firstBottom)))),
            CellLayoutStyle(31, BytesField(9, Message(FloatField(2, 3), FloatField(4, secondBottom)))),
            automatic ? TableSizingStyle(40, true) : Array.Empty<byte>()
        }, modelDeclarations: Message(DoubleField(16, height), automatic ? ReferenceField(3, 40) : Array.Empty<byte>()), secondStyleReference: 31,
            merge: merge == "horizontal", rowCount: rowCount,
            mergeDeclarations: merge.StartsWith("vertical", StringComparison.Ordinal)
                ? MergeTestOwner(MergeTestPair(0, 0, 1, 0)) : null,
            additionalTileRows: merge == "vertical-styled" ? PhysicalContinuationRow() : null);

    private static byte[] PhysicalContinuationRow() {
        byte[] empty = new byte[16]; empty[0] = 5; empty[8] = 0x20; empty[12] = 2;
        return BytesField(5, Message(VarintField(1, 1), VarintField(2, 1),
            BytesField(6, empty), BytesField(7, new byte[] { 0, 0 })));
    }
}
