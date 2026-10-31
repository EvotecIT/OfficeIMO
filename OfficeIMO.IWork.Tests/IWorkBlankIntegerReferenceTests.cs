using OfficeIMO.Excel;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Integer_references_admit_complete_sparse_and_materialized_blanks_beside_numbers(
        bool materialized, bool scalar) {
        var cells = materialized
            ? new[] { new IWorkTableCell(1, 1, IWorkCellKind.Empty, null) }
            : Array.Empty<IWorkTableCell>();
        var table = new IWorkTable("Inputs", 1, 1, cells);
        var own = new IWorkFormulaTableBinding(table, string.Empty, false);
        foreach (int function in new[] { 58, 75 }) {
            var formula = IntegerReferenceFormula(function,
                scalar ? IntegerTextSingleCell() : IntegerRectangle(0),
                Message(VarintField(1, 17), DoubleField(4, 6)));
            var result = IWorkFormulaReader.Render(formula, 0, 0, 100, 1000, owningTable: own,
                budget: new IWorkProjectionBudget(new IWorkReadOptions()));
            Assert.True(result.IsComplete);
            string reference = scalar ? "$A$1" : "$A$1:$A$1";
            Assert.Equal("=" + (function == 58 ? "GCD" : "LCM") + "(" + reference + ",6)", result.Text);
            using var destination = ExcelDocument.Create();
            var sheet = destination.AddSheet("Inputs");
            sheet.CellFormula(1, 2, result.Text);
            using var bytes = new MemoryStream(); destination.Save(bytes); bytes.Position = 0;
            using var reopened = ExcelDocument.Load(bytes);
            Assert.Equal(1, reopened.Calculate());
            Assert.Equal(function == 58 || scalar ? 6d : 0d, reopened.Sheets[0].CellAt(1, 2).GetValue<double>());
        }
    }

    [Fact]
    public void All_blank_integer_arguments_keep_the_native_unqualified_error_boundary() {
        var own = new IWorkFormulaTableBinding(new IWorkTable("Inputs", 1, 1,
            Array.Empty<IWorkTableCell>()), string.Empty, false);
        foreach (int function in new[] { 58, 75 }) {
            foreach (bool scalar in new[] { false, true }) {
                var formula = IntegerReferenceFormula(function,
                    scalar ? IntegerTextSingleCell() : IntegerRectangle(0));
                Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 100, 1000, owningTable: own,
                    budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
            }
        }
    }

    [Theory]
    [InlineData("missingTile")]
    [InlineData("malformedRow")]
    [InlineData("rowCount")]
    [InlineData("invalidOffset")]
    [InlineData("legacy")]
    [InlineData("duplicateRow")]
    [InlineData("declaredRowCount")]
    public void Unreadable_sparse_storage_cannot_turn_absent_cells_into_integer_zeroes(string failure) {
        byte[] row = Message(VarintField(1, 0), BytesField(6, Array.Empty<byte>()),
            BytesField(7, new byte[] { 255, 255 }));
        byte[] tile = failure switch {
            "malformedRow" => BytesField(5, new byte[] { 0 }),
            "rowCount" => BytesField(5, Message(row, VarintField(2, 1))),
            "invalidOffset" => BytesField(5, Message(VarintField(1, 0), BytesField(6, new byte[] { 5 }),
                BytesField(7, new byte[] { 254, 0 }))),
            "legacy" => BytesField(5, Message(VarintField(1, 0), BytesField(3, new byte[] { 5 }),
                BytesField(4, new byte[] { 0, 0 }))),
            "duplicateRow" => Message(BytesField(5, row), BytesField(5, row)),
            "declaredRowCount" => VarintField(4, 1),
            _ => BytesField(5, row)
        };
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(),
            includeTile: failure != "missingTile", tilePayload: tile);
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        if (failure == "invalidOffset") Assert.True(table.GetCell(1, 1)!.HasDecodeError);
        else Assert.Null(table.GetCell(1, 1));
        Assert.False(RenderBlankInteger(table, scalar: true).IsComplete);
        Assert.False(RenderBlankInteger(table, scalar: false).IsComplete);
    }

    [Theory]
    [InlineData(0UL)]
    [InlineData(2UL)]
    public void Inconsistent_tile_row_counts_retain_readable_siblings_without_establishing_blanks(ulong declaredRows) {
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), rows: 2,
            tilePayload: Message(VarintField(4, declaredRows), BytesField(5, TileTestRow(1))));
        var source = IWorkSourceDocument.Open(package);
        var projection = source.ReadNumbers();
        var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Equal(42d, Assert.Single(table.Cells).Value);
        Assert.Equal(2, Assert.Single(table.Cells).Row);
        Assert.Null(table.GetCell(1, 1));
        Assert.False(projection.HasEditableContent);
        Assert.False(RenderBlankInteger(table, scalar: true).IsComplete);
        var issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal(12UL, issue.Owner.RecordIdentifier);
        Assert.Equal("4", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, issue.Kind);
        Assert.Throws<NotSupportedException>(() => source.BeginEdit().SetNumbersCellNumber(table.SourceIdentity!, 2, 1, 43));
    }

    [Theory]
    [InlineData("different")]
    [InlineData("zero")]
    [InlineData("duplicate")]
    [InlineData("wrongWire")]
    public void Ambiguous_tile_sizes_cannot_relocate_cells_or_establish_blanks(string failure) {
        byte[] size = failure switch {
            "different" => VarintField(2, 128),
            "zero" => VarintField(2, 0),
            "duplicate" => Message(VarintField(2, 256), VarintField(2, 256)),
            _ => BytesField(2, Message())
        };
        byte[] tileStorage = Message(size, BytesField(1, Message(VarintField(1, 1), ReferenceField(2, 12))));
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, BytesField(3, tileStorage),
            includeTile: false, rows: 257, tilePayload: BytesField(5, TileTestRow(0)));
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Empty(table.Cells);
        Assert.False(projection.HasEditableContent);
        Assert.False(RenderBlankInteger(table, scalar: true).IsComplete);
        var issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal(11UL, issue.Owner.RecordIdentifier);
        Assert.Equal("4/3/2", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, issue.Kind);
    }

    [Fact]
    public void Qualified_native_tile_size_and_row_count_keep_zero_legacy_cell_counts_and_sparse_positions() {
        byte[] tileStorage = Message(VarintField(2, 256),
            BytesField(1, Message(VarintField(1, 1), ReferenceField(2, 12))));
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, BytesField(3, tileStorage),
            includeTile: false, rows: 257, tilePayload: Message(VarintField(3, 0), VarintField(4, 1),
                BytesField(5, TileTestRow(0))));
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Equal(257, Assert.Single(table.Cells).Row);
        Assert.Equal(42d, Assert.Single(table.Cells).Value);
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.SourceDeclarationIssues);
        Assert.True(RenderBlankInteger(table, scalar: true).IsComplete);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Valid_native_sparse_storage_establishes_blanks_without_materializing_them(bool stored) {
        byte[] empty = stored ? new byte[12] : Array.Empty<byte>();
        if (stored) empty[0] = 5;
        byte[] row = Message(VarintField(1, 0), VarintField(2, stored ? 1UL : 0UL), BytesField(6, empty),
            BytesField(7, stored ? new byte[] { 0, 0 } : new byte[] { 255, 255 }));
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), tilePayload: BytesField(5, row));
        var table = Assert.Single(Assert.Single(IWorkSourceDocument.Open(package).ReadNumbers().Sheets).Tables);
        Assert.Empty(table.Cells);
        Assert.True(RenderBlankInteger(table, scalar: true).IsComplete);
        Assert.True(RenderBlankInteger(table, scalar: false).IsComplete);
        Assert.True(RenderBlankInteger(table.WithCells(table.Cells), scalar: true).IsComplete);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Blank_integer_references_require_complete_metadata_and_consistent_empty_values(bool materialized) {
        var cells = materialized ? new[] { new IWorkTableCell(1, 1, IWorkCellKind.Empty, null) }
            : Array.Empty<IWorkTableCell>();
        var badBody = new IWorkTable("Inputs", 1, 1, cells, bodyMetadataIsComplete: false);
        var badStorage = new IWorkTable("Inputs", 1, 1, cells, cellStorageIsComplete: false);
        foreach (var table in new[] { badBody, badStorage, badStorage.WithCells(cells) })
            Assert.False(RenderBlankInteger(table, scalar: true).IsComplete);
        foreach (var cell in new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Empty, null, cachedValueIsComplete: false),
            new IWorkTableCell(1, 1, IWorkCellKind.Empty, null, sourceFormulaIsDeclared: true),
            new IWorkTableCell(1, 1, IWorkCellKind.Empty, null, hasDecodeError: true),
            new IWorkTableCell(1, 1, IWorkCellKind.Empty, null, valueKind: IWorkCellKind.Number),
            new IWorkTableCell(1, 1, IWorkCellKind.Empty, "12")
        }) Assert.False(RenderBlankInteger(new IWorkTable("Inputs", 1, 1, new[] { cell }), scalar: true).IsComplete);
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void Only_range_blanks_supply_LCM_zero_when_other_products_exceed_the_native_bound(bool scalar, bool complete) {
        var table = new IWorkTable("Inputs", 1, 1, Array.Empty<IWorkTableCell>());
        var formula = IntegerReferenceFormula(75, scalar ? IntegerTextSingleCell() : IntegerRectangle(0),
            Message(VarintField(1, 17), DoubleField(4, 999999999999999d)),
            Message(VarintField(1, 17), DoubleField(4, 2)));
        Assert.Equal(complete, IWorkFormulaReader.Render(formula, 0, 0, 100, 1000,
            owningTable: new IWorkFormulaTableBinding(table, string.Empty, false),
            budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
    }

    private static IWorkFormulaResult RenderBlankInteger(IWorkTable table, bool scalar) =>
        IWorkFormulaReader.Render(IntegerReferenceFormula(75,
            scalar ? IntegerTextSingleCell() : IntegerRectangle(0),
            Message(VarintField(1, 17), DoubleField(4, 6))), 0, 0, 100, 1000,
            owningTable: new IWorkFormulaTableBinding(table, string.Empty, false),
            budget: new IWorkProjectionBudget(new IWorkReadOptions()));
}
