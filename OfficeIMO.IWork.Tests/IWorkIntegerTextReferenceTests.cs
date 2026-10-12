using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_integer_text_single_cells_keep_type_caches_and_saved_operand_recalculation(bool changed) {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture(
            "native-exports/numbers-integer-text-reference-v14.5-decode.json")));
        JsonElement evidence = manifest.RootElement;
        foreach (JsonElement input in evidence.GetProperty("inputs").EnumerateArray()) {
            byte[] bytes = File.ReadAllBytes(CorpusFixture("native-exports/" + input.GetProperty("path").GetString()));
            Assert.Equal(input.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(CorpusFixture("native-exports/" +
            evidence.GetProperty(changed ? "changedSource" : "source").GetString()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        IWorkTable operands = result.Projection.Sheets[0].Tables.Single(), checks = result.Projection.Sheets[1].Tables.Single();
        Assert.Equal(IWorkCellKind.Text, operands.GetCell(4, 2)!.Kind);
        Assert.Equal(changed ? "18" : "12", operands.GetCell(4, 2)!.Value);
        Assert.Equal(4, checks.Cells.Count(c => c.Kind == IWorkCellKind.Formula && c.FormulaIsComplete));
        // Blank references, mixed ranges, Booleans and arbitrary text remain
        // outside this feature even when native Numbers has a numeric cache.
        foreach (int row in new[] { 2, 3, 5, 6, 8 }) {
            Assert.False(checks.GetCell(row, 2)!.FormulaIsComplete);
            Assert.False(checks.GetCell(row, 3)!.FormulaIsComplete);
        }
        using var bytesOut = new MemoryStream(); result.Value.Save(bytesOut); bytesOut.Position = 0;
        using var reopened = ExcelDocument.Load(bytesOut);
        Assert.Equal(new[] { "Inputs", "Checks" }, reopened.Sheets.Select(s => s.Name));
        Assert.Equal(changed ? "18" : "12", reopened.Sheets[0].CellAt(4, 2).GetValue<string>());
        Verify(changed ? "changedCells" : "cells", true);
        Assert.Equal(4, reopened.Calculate()); Verify(changed ? "changedCells" : "cells", false);
        reopened.Sheets[0].CellValue(4, 2, "18");
        Assert.Equal(4, reopened.Calculate()); Verify("changedCells", false);
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes);
        Assert.Equal("18", final.Sheets[0].CellAt(4, 2).GetValue<string>());
        Assert.Equal(18d, final.Sheets[1].CellAt(4, 3).GetValue<double>());
        Assert.Equal(18d, final.Sheets[1].CellAt(7, 2).GetValue<double>());
        Assert.Empty(final.ValidateOpenXml());

        void Verify(string field, bool sourceToo) {
            foreach (JsonElement expected in evidence.GetProperty(field).EnumerateArray()) {
                int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
                if (sourceToo) {
                    IWorkTableCell cell = checks.GetCell(row, column)!;
                    Assert.True(cell.FormulaIsComplete); Assert.True(cell.CachedValueIsComplete);
                    Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), cell.Value);
                }
                Assert.Equal(expected.GetProperty("excelFormula").GetString(), reopened.Sheets[1].GetFormulaText(row, column));
                Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), reopened.Sheets[1].CellAt(row, column).GetValue<double>());
            }
        }
    }

    [Theory]
    [InlineData("0", true)]
    [InlineData("12", true)]
    [InlineData("999999999999999", true)]
    [InlineData("1000000000000000", false)]
    [InlineData("01", false)]
    [InlineData(" 12", false)]
    [InlineData("12.5", false)]
    [InlineData("-1", false)]
    [InlineData("١٢", false)]
    public void Integer_text_references_admit_only_canonical_nonnegative_integers(string text, bool complete) {
        var table = new IWorkTable("Inputs", 1, 1, new[] { new IWorkTableCell(1, 1, IWorkCellKind.Text, text) });
        var own = new IWorkFormulaTableBinding(table, string.Empty, false);
        foreach (int function in new[] { 58, 75 }) {
            var formula = IntegerReferenceFormula(function, IntegerTextSingleCell());
            Assert.Equal(complete, IWorkFormulaReader.Render(formula, 0, 0, 100, 1000,
                owningTable: own, budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
        }
    }

    [Fact]
    public void Integer_text_references_keep_ranges_and_incomplete_scalar_metadata_partial_even_beside_zero() {
        var cells = new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Text, "12", cachedValueIsComplete: false),
            new IWorkTableCell(1, 1, IWorkCellKind.Text, "12", hasDecodeError: true),
            new IWorkTableCell(1, 1, IWorkCellKind.Text, "12", sourceFormulaIsDeclared: true),
            new IWorkTableCell(1, 1, IWorkCellKind.Text, "12", valueKind: IWorkCellKind.Number)
        };
        foreach (IWorkTableCell cell in cells) {
            var own = new IWorkFormulaTableBinding(new IWorkTable("Inputs", 1, 1, new[] { cell }), string.Empty, false);
            var formula = IntegerReferenceFormula(75, Message(VarintField(1, 17), DoubleField(4, 0)), IntegerTextSingleCell());
            Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 100, 1000, owningTable: own,
                budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
        }
        var text = new IWorkTableCell(1, 1, IWorkCellKind.Text, "12");
        var single = new IWorkFormulaTableBinding(new IWorkTable("Inputs", 1, 1, new[] { text }), string.Empty, false);
        var rectangle = IntegerReferenceFormula(75, IntegerRectangle(0));
        Assert.False(IWorkFormulaReader.Render(rectangle, 0, 0, 100, 1000, owningTable: single,
            budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
        // A source colon expression remains a range even when both endpoints
        // identify the same cell. Native-normalized single cells use type 36.
        var colon = IWorkProtobuf.Parse(BytesField(1, Message(BytesField(1, IntegerTextSingleCell()),
            BytesField(1, IntegerTextSingleCell()), BytesField(1, VarintField(1, 29)),
            BytesField(1, Message(VarintField(1, 16), VarintField(2, 75), VarintField(3, 1))))), new IWorkReadOptions());
        Assert.False(IWorkFormulaReader.Render(colon, 0, 0, 100, 1000, owningTable: single,
            budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
    }

    [Theory]
    [InlineData("1\uFFFC2", false)]
    [InlineData("1\uFFFB2", false)]
    [InlineData("12", true)]
    public void Integer_text_references_require_complete_decoded_rich_text(string nativeText, bool complete) {
        using var package = CreateNumbersWithPartialRichCell(hasFormula: false, richText: nativeText);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        IWorkTable table = result.Projection.Sheets.Single().Tables.Single();
        IWorkTableCell source = table.Cells.Single();
        Assert.Equal("12", source.Value);
        Assert.Equal(complete, source.RichText!.IsTextComplete);
        var own = new IWorkFormulaTableBinding(table, string.Empty, false);
        foreach (int function in new[] { 58, 75 }) {
            var formula = IntegerReferenceFormula(function, IntegerTextSingleCell());
            Assert.Equal(complete, IWorkFormulaReader.Render(formula, 0, 0, 100, 1000,
                owningTable: own, budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
        }
        Assert.Equal(complete, source.CachedValueIsComplete);
    }

    private static byte[] IntegerTextSingleCell() => Message(VarintField(1, 36),
        BytesField(26, Message(VarintField(1, 0), VarintField(2, 1))),
        BytesField(27, Message(VarintField(1, 0), VarintField(2, 1))));
}
