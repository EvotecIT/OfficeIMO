using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Region_layouts_reach_saved_destinations_without_materializing_empty_source_cells(IWorkDocumentKind kind) {
        using var package = CellFillPackage(new[] {
            CellLayoutStyle(30, BytesField(9, FloatField(1, 2))),
            CellLayoutStyle(31, Message(BytesField(9, Message()), VarintField(8, 0))),
            RegionLayout(41, 4, 1), RegionLayout(42, 5, 2), RegionLayout(43, 6, 1), RegionLayout(44, 8, 2)
        }, kind: kind, rowCount: 4, secondStyleReference: 31,
            modelDeclarations: Message(VarintField(9, 1), VarintField(10, 1), VarintField(11, 1), DoubleField(16, 40),
                ReferenceField(18, 41), ReferenceField(19, 42), ReferenceField(20, 43), ReferenceField(21, 44)));
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        IWorkTable table;
        (int Row, int Column, double Padding, IWorkCellVerticalAlignment Vertical)[] defaults = {
            (2, 1, 6, IWorkCellVerticalAlignment.Middle), (2, 2, 4, IWorkCellVerticalAlignment.Middle),
            (4, 1, 8, IWorkCellVerticalAlignment.Bottom)
        };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy); table = Assert.Single(result.Projection.Tables);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = WordDocument.Load(saved);
            WordTable target = Assert.Single(reopened.Tables);
            foreach (var expected in defaults) {
                var cell = target.Rows[expected.Row - 1].Cells[expected.Column - 1];
                Assert.Equal((short)(expected.Padding * 20), checked((short?)(cell.MarginBottom?.ToTwips())));
                Assert.Equal(expected.Vertical == IWorkCellVerticalAlignment.Middle ? WordTableVerticalAlignment.Center : WordTableVerticalAlignment.Bottom,
                    cell.VerticalAlignment);
            }
            Assert.Equal(new[] { 800, 680, 680, 640 }, target.RowHeight);
            Assert.Equal((short)40, checked((short?)(target.Rows[0].Cells[0].MarginLeft?.ToTwips())));
            Assert.Equal((short)0, checked((short?)(target.Rows[0].Cells[0].MarginBottom?.ToTwips())));
            Assert.Equal(WordTableVerticalAlignment.Bottom, target.Rows[0].Cells[0].VerticalAlignment);
            Assert.Equal(WordTableVerticalAlignment.Top, target.Rows[0].Cells[1].VerticalAlignment);
            Assert.Empty(reopened.ValidateDocument());
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(policy); table = Assert.Single(result.Projection.Sheets[0].Tables);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED");
            Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal("center", reopened.Sheets[0].GetCellStyle(2, 2).VerticalAlignment);
            Assert.Equal("bottom", reopened.Sheets[0].GetCellStyle(4, 1).VerticalAlignment);
            Assert.Equal("top", reopened.Sheets[0].GetCellStyle(1, 2).VerticalAlignment);
        } else {
            using var result = source.ToPowerPointDocumentResult(policy); table = Assert.Single(result.Projection.Slides[0].Tables);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointDocument.Load(saved);
            var target = Assert.Single(reopened.Slides[0].Tables);
            foreach (var expected in defaults) {
                var cell = target.GetCell(expected.Row - 1, expected.Column - 1);
                Assert.Equal(expected.Padding, cell.PaddingBottomPoints);
                Assert.Equal(expected.Vertical == IWorkCellVerticalAlignment.Middle ? PowerPointTextVerticalAlignment.Center : PowerPointTextVerticalAlignment.Bottom,
                    cell.VerticalAlignment);
            }
            Assert.Equal(0, target.GetCell(0, 0).PaddingBottomPoints);
            Assert.Equal(PowerPointTextVerticalAlignment.Bottom, target.GetCell(0, 0).VerticalAlignment);
            Assert.Equal(PowerPointTextVerticalAlignment.Top, target.GetCell(0, 1).VerticalAlignment);
            Assert.Empty(reopened.ValidateDocument());
        }
        Assert.Equal(2, table.Cells.Count); Assert.Null(table.GetCell(2, 2));
        Assert.Equal(2, table.GetPadding(1, 1)!.LeftPoints); Assert.Equal(0, table.GetPadding(1, 1)!.BottomPoints);
        Assert.Equal(0, table.GetPadding(1, 2)!.LeftPoints);
        foreach (var expected in defaults) {
            Assert.Equal(expected.Padding, table.GetPadding(expected.Row, expected.Column)!.BottomPoints);
            Assert.Equal(expected.Vertical, table.GetVerticalAlignment(expected.Row, expected.Column));
        }
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetPadding(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetVerticalAlignment(1, 3));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Invalid_region_or_selected_layout_does_not_fall_back_and_preserves_the_other_property(bool selected, bool alignment) {
        byte[] invalid = alignment ? Message(VarintField(8, 3), BytesField(9, CellPadding(2)))
            : Message(VarintField(8, 0), BytesField(9, CellPadding(-1)));
        using var package = CellFillPackage(new[] {
            CellLayoutStyle(30, selected ? invalid : Message()),
            CellLayoutStyle(40, selected ? Message(VarintField(8, 1), BytesField(9, CellPadding(4))) : invalid)
        }, rowCount: 2, modelDeclarations: ReferenceField(18, 40));
        var projection = IWorkSourceDocument.Open(package).ReadPages();
        var table = Assert.Single(projection.Tables);
        Assert.False(projection.HasEditableContent); Assert.Equal(42d, table.GetCell(1, 1)!.Value);
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == (selected
            ? "IWORK_TABLE_CELL_LAYOUT_UNSUPPORTED" : "IWORK_TABLE_CELL_LAYOUT_DEFAULTS_UNSUPPORTED"));
        int row = selected ? 1 : 2;
        if (alignment) {
            Assert.Null(table.GetVerticalAlignment(row, 1));
            Assert.Equal(2, table.GetPadding(row, 1)!.BottomPoints);
        } else {
            Assert.Null(table.GetPadding(row, 1));
            Assert.Equal(IWorkCellVerticalAlignment.Top, table.GetVerticalAlignment(row, 1));
        }
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.Owner.RecordIdentifier == (selected ? 30ul : 40ul));
    }

    [Fact]
    public void Unused_region_layout_references_are_not_traversed_and_inherited_layout_reaches_unstored_cells() {
        using var package = CellFillPackage(new[] { CellLayoutStyle(30, Message()), CellLayoutStyle(40, Message(), 41), RegionLayout(41, 4, 1) },
            rowCount: 2, modelDeclarations: Message(ReferenceField(18, 40), ReferenceField(19, 99)));
        var projection = IWorkSourceDocument.Open(package).ReadPages(); var table = Assert.Single(projection.Tables);
        Assert.True(projection.HasEditableContent); Assert.Empty(projection.SourceReferenceIssues); Assert.Empty(projection.SourceDeclarationIssues);
        Assert.Null(table.GetCell(2, 2)); Assert.Equal(4, table.GetPadding(2, 2)!.LeftPoints);
        Assert.Equal(IWorkCellVerticalAlignment.Middle, table.GetVerticalAlignment(2, 2));
    }

    private static byte[] RegionLayout(ulong id, float padding, ulong vertical) =>
        CellLayoutStyle(id, Message(BytesField(9, CellPadding(padding)), VarintField(8, vertical)));
}
