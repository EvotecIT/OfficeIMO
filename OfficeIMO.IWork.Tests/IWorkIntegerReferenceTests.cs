using System.Threading;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Apple_edited_integer_references_retain_native_caches_and_recalculate_after_saved_operand_edits() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture(
            "native-exports/numbers-integer-references-v14.5-decode.json")));
        JsonElement evidence = manifest.RootElement;
        foreach (JsonElement input in evidence.GetProperty("inputs").EnumerateArray()) {
            string path = CorpusFixture("native-exports/" + input.GetProperty("path").GetString());
            Assert.Equal(input.GetProperty("sha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        }
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(CorpusFixture(
            "native-exports/" + evidence.GetProperty("source").GetString()),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        IWorkTable table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
        Assert.Equal(33, table.Cells.Count(c => c.Kind == IWorkCellKind.Formula && c.FormulaIsComplete));
        Assert.False(table.GetCell(10, 6)!.FormulaIsComplete);
        Assert.Equal(0d, table.GetCell(10, 6)!.Value);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);
        foreach (JsonElement expected in evidence.GetProperty("cells").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
            Assert.True(table.GetCell(row, column)!.FormulaIsComplete);
            Assert.Equal(expected.GetProperty("formula").GetString(), sheet.GetFormulaText(row, column));
            Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), sheet.CellAt(row, column).GetValue<double>());
        }
        Assert.Equal(33, reopened.Calculate());
        sheet.CellValue(2, 2, 24d); Assert.Equal(33, reopened.Calculate());
        foreach (int column in new[] { 5, 7 }) Assert.Equal(12d, sheet.CellAt(2, column).GetValue<double>());
        foreach (int column in new[] { 6, 8 }) Assert.Equal(360d, sheet.CellAt(2, column).GetValue<double>());
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes);
        Assert.Equal("GCD(B2,C2,D2)", final.Sheets[0].GetFormulaText(2, 7));
        Assert.Equal("LCM($B$2,$C$2:$D$2,1)", final.Sheets[0].GetFormulaText(2, 8));
        Assert.Equal(12d, final.Sheets[0].CellAt(2, 7).GetValue<double>());
        Assert.Equal(360d, final.Sheets[0].CellAt(2, 8).GetValue<double>());
        Assert.Empty(final.ValidateOpenXml());
    }

    [Fact]
    public void Integer_references_use_canonical_identity_and_retain_endpoint_and_literal_operands() {
        Guid id = Guid.Parse("00112233-4455-6677-8899-aabbccddeeff");
        var table = new IWorkTable("Inputs", 1, 3, new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Number, 24d),
            new IWorkTableCell(1, 2, IWorkCellKind.Number, 36d),
            new IWorkTableCell(1, 3, IWorkCellKind.Number, 60d)
        }, formulaIdentifier: id);
        var own = new IWorkFormulaTableBinding(table, string.Empty, false);
        byte[] external = BytesField(28, BytesField(1, Message(VarintField(2, 0x33221100),
            VarintField(3, 0x77665544), VarintField(4, 0xbbaa9988), VarintField(5, 0xffeeddcc))));
        var bindings = IWorkFormulaTableBindings.Create(new[] { table }, _ => "'Inputs'!", default);
        IWorkWireMessage formula = IntegerReferenceFormula(58, IntegerRectangle(2, external),
            Message(VarintField(1, 17), DoubleField(4, 6)));
        IWorkFormulaResult result = IWorkFormulaReader.Render(formula, 0, 0, 100, 1000, bindings,
            budget: new IWorkProjectionBudget(new IWorkReadOptions()));
        Assert.True(result.IsComplete);
        Assert.Equal("=GCD('Inputs'!$A$1:$C$1,6)", result.Text);
        Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 100, 1000, owningTable: own,
            budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
        var duplicate = IWorkFormulaTableBindings.Create(new[] { table, table }, _ => "'Inputs'!", default);
        Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 100, 1000, duplicate,
            budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);

        byte[] coordinate(int column) => Message(VarintField(1, 36),
            BytesField(26, Message(VarintField(1, (ulong)column * 2), VarintField(2, 1))),
            BytesField(27, Message(VarintField(1, 0), VarintField(2, 1))));
        IWorkWireMessage endpoints = IWorkProtobuf.Parse(BytesField(1, Message(
            BytesField(1, coordinate(0)), BytesField(1, coordinate(2)), BytesField(1, VarintField(1, 29)),
            BytesField(1, Message(VarintField(1, 16), VarintField(2, 75), VarintField(3, 1))))), new IWorkReadOptions());
        result = IWorkFormulaReader.Render(endpoints, 0, 0, 100, 1000, owningTable: own,
            budget: new IWorkProjectionBudget(new IWorkReadOptions()));
        Assert.True(result.IsComplete);
        Assert.Equal("=LCM($A$1:$C$1)", result.Text);
    }

    [Fact]
    public void Integer_reference_qualification_rejects_unqualified_cells_even_beside_literal_zero() {
        IWorkTableCell[] unqualified = {
            new(1, 1, IWorkCellKind.Empty, null),
            new(1, 1, IWorkCellKind.Text, "24"),
            new(1, 1, IWorkCellKind.Boolean, true),
            new(1, 1, IWorkCellKind.Formula, 24d, formula: "=12*2", valueKind: IWorkCellKind.Number, formulaIsComplete: true),
            new(1, 1, IWorkCellKind.Number, -0.2d),
            new(1, 1, IWorkCellKind.Number, 1000000000000000d),
            new(1, 1, IWorkCellKind.Number, 24d, numericValueIsApproximate: true),
            new(1, 1, IWorkCellKind.Number, 24d, cachedValueIsComplete: false),
            new(1, 1, IWorkCellKind.Number, 24d, hasDecodeError: true),
            new(1, 1, IWorkCellKind.Error, "#ERROR")
        };
        IWorkWireMessage formula = IntegerReferenceFormula(75,
            Message(VarintField(1, 17), DoubleField(4, 0)), IntegerRectangle(0));
        foreach (IWorkTableCell cell in unqualified) {
            var table = new IWorkTable("Inputs", 1, 1, new[] { cell });
            Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 100, 1000,
                owningTable: new IWorkFormulaTableBinding(table, string.Empty, false),
                budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
        }
        var missing = new IWorkTable("Inputs", 1, 1, Array.Empty<IWorkTableCell>());
        Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 100, 1000,
            owningTable: new IWorkFormulaTableBinding(missing, string.Empty, false),
            budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
        var merged = new IWorkTable("Inputs", 1, 2, new[] { new IWorkTableCell(1, 1, IWorkCellKind.Number, 24d) },
            mergedRanges: new[] { new IWorkTableMergeRange(1, 1, 1, 2) });
        Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 100, 1000,
            owningTable: new IWorkFormulaTableBinding(merged, string.Empty, false),
            budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
    }

    [Fact]
    public void Integer_reference_inspection_obeys_formula_position_and_source_work_limits_and_cancellation() {
        var cells = Enumerable.Range(1, 5).Select(c => new IWorkTableCell(1, c, IWorkCellKind.Number, 12d)).ToArray();
        var own = new IWorkFormulaTableBinding(new IWorkTable("Inputs", 1, 5, cells), string.Empty, false);
        IWorkWireMessage formula = IntegerReferenceFormula(58, IntegerRectangle(4));
        var options = new IWorkReadOptions { MaximumFormulaRenderingOperations = 5 };
        Assert.False(IWorkFormulaReader.Render(formula, 0, 0, 4, 1000, owningTable: own,
            budget: new IWorkProjectionBudget(options)).IsComplete);
        var budget = new IWorkProjectionBudget(options);
        Assert.True(IWorkFormulaReader.Render(formula, 0, 0, 5, 1000, owningTable: own, budget: budget).IsComplete);
        Assert.Throws<InvalidDataException>(() => IWorkFormulaReader.Render(formula, 0, 0, 5, 1000,
            owningTable: own, budget: budget));
        Assert.Throws<OperationCanceledException>(() => IWorkFormulaReader.Render(formula, 0, 0, 5, 1000,
            owningTable: own, budget: new IWorkProjectionBudget(options), cancellationToken: new CancellationToken(true)));
        IWorkWireMessage repeated = IntegerReferenceFormula(58, IntegerRectangle(4), IntegerRectangle(4));
        Assert.False(IWorkFormulaReader.Render(repeated, 0, 0, 5, 1000, owningTable: own,
            budget: new IWorkProjectionBudget(new IWorkReadOptions())).IsComplete);
    }

    private static byte[] IntegerRectangle(int lastColumn, byte[]? identity = null) => Message(VarintField(1, 67),
        BytesField(40, Message(BytesField(3, Message(VarintField(1, 0), VarintField(2, (ulong)lastColumn))),
            BytesField(4, Message(VarintField(1, 0), VarintField(2, 0))))), identity ?? Array.Empty<byte>());

    private static IWorkWireMessage IntegerReferenceFormula(int function, params byte[][] operands) =>
        IWorkProtobuf.Parse(BytesField(1, Message(operands.Select(o => BytesField(1, o)).Append(
            BytesField(1, Message(VarintField(1, 16), VarintField(2, (ulong)function),
                VarintField(3, (ulong)operands.Length)))).ToArray())), new IWorkReadOptions());
}
