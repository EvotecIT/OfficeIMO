using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_blank_integer_references_preserve_scalar_range_semantics_and_saved_operand_edits(bool filled) {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture(
            "native-exports/numbers-blank-reference-v14.5-decode.json")));
        JsonElement evidence = manifest.RootElement;
        VerifyInputs();
        JsonElement native = evidence.GetProperty("cases")[filled ? 1 : 0];
        JsonElement edited = evidence.GetProperty("cases")[1];
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(CorpusFixture(
            "native-exports/" + native.GetProperty("source").GetString()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        Assert.True(result.Report.IsPartialEditableReconstruction);
        IWorkTable operands = result.Projection.Sheets[0].Tables.Single();
        IWorkTable checks = result.Projection.Sheets[1].Tables.Single();
        Assert.Equal("Inputs - Operands", result.Projection.Sheets[0].Name);
        Assert.Equal("Checks - Results", result.Projection.Sheets[1].Name);
        Assert.Equal(IWorkCellKind.Text, operands.GetCell(5, 3)!.Kind);
        Assert.Equal("12", operands.GetCell(5, 3)!.Value);
        if (filled) Assert.Equal(18d, operands.GetCell(4, 3)!.Value);
        else Assert.Null(operands.GetCell(4, 3));
        int qualified = native.GetProperty("qualifiedFormulaCells").GetInt32();
        Assert.Equal(14, checks.Cells.Count(c => c.Kind == IWorkCellKind.Formula));
        Assert.Equal(qualified, checks.Cells.Count(c => c.Kind == IWorkCellKind.Formula && c.FormulaIsComplete));
        using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        Assert.Equal(new[] { "Inputs - Operands", "Checks - Results" }, reopened.Sheets.Select(s => s.Name));
        VerifyCells(native, sourceToo: true, afterOperandEdit: false);
        Assert.Equal(qualified, reopened.Calculate());
        VerifyCells(native, sourceToo: false, afterOperandEdit: false);
        reopened.Sheets[0].CellValue(4, 3, 18d);
        Assert.Equal(qualified, reopened.Calculate());
        VerifyCells(edited, sourceToo: false, afterOperandEdit: true);
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes);
        Assert.Equal(18d, final.Sheets[0].CellAt(4, 3).GetValue<double>());
        Assert.Equal(18d, final.Sheets[1].CellAt(3, 4).GetValue<double>());
        Assert.Equal(90d, final.Sheets[1].CellAt(4, 4).GetValue<double>());
        if (!filled) {
            Assert.Null(final.Sheets[1].GetFormulaText(6, 4));
            Assert.Equal("#ERROR", final.Sheets[1].CellAt(6, 4).GetValue<string>());
        }
        Assert.Empty(final.ValidateOpenXml());
        VerifyInputs();

        void VerifyCells(JsonElement expectedCase, bool sourceToo, bool afterOperandEdit) {
            foreach (JsonElement expected in expectedCase.GetProperty("cells").EnumerateArray()) {
                int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
                bool isEditable = filled || row is not (6 or 7);
                IWorkTableCell source = checks.GetCell(row, column)!;
                if (sourceToo) {
                    Assert.Equal(isEditable, source.FormulaIsComplete);
                    Assert.True(source.CachedValueIsComplete);
                    Assert.Equal(isEditable ? IWorkCellKind.Number : IWorkCellKind.Error, source.ValueKind);
                    if (isEditable) Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), source.Value);
                    else Assert.Equal("#ERROR", source.Value);
                }
                if (isEditable) {
                    Assert.Equal(expected.GetProperty("excelFormula").GetString(), reopened.Sheets[1].GetFormulaText(row, column));
                    Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), reopened.Sheets[1].CellAt(row, column).GetValue<double>());
                } else {
                    // An omitted source formula remains a source error cache
                    // when a later destination edit changes its former operand.
                    Assert.Null(reopened.Sheets[1].GetFormulaText(row, column));
                    Assert.Equal("#ERROR", reopened.Sheets[1].CellAt(row, column).GetValue<string>());
                    if (!afterOperandEdit) Assert.False(expected.GetProperty("qualified").GetBoolean());
                }
            }
        }

        void VerifyInputs() {
            foreach (JsonElement input in evidence.GetProperty("inputs").EnumerateArray()) {
                byte[] inputBytes = File.ReadAllBytes(CorpusFixture("native-exports/" + input.GetProperty("path").GetString()));
                Assert.Equal(input.GetProperty("bytes").GetInt64(), inputBytes.LongLength);
                Assert.Equal(input.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(inputBytes)).ToLowerInvariant());
            }
        }
    }
}
