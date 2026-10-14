using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_conditional_aggregates_retain_formulas_and_match_fresh_Excel_before_and_after_saved_edits() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture(
            "native-exports/numbers-conditional-aggregates-v14.5-decode.json")));
        JsonElement evidence = manifest.RootElement;
        foreach (JsonElement input in evidence.GetProperty("inputs").EnumerateArray()) {
            byte[] bytes = File.ReadAllBytes(CorpusFixture("native-exports/" + input.GetProperty("path").GetString()));
            Assert.Equal(input.GetProperty("bytes").GetInt64(), bytes.LongLength);
            Assert.Equal(input.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(CorpusFixture(
            "native-exports/" + evidence.GetProperty("source").GetString()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED");
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        var table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
        Assert.Equal(20, table.Cells.Count(c => c.Kind == IWorkCellKind.Formula));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var sheet = Assert.Single(reopened.Sheets);
        Verify(sheet, edited: false, sourceToo: true);
        Assert.Equal(20, reopened.Calculate());
        Verify(sheet, edited: false, sourceToo: false);
        sheet.CellValue(2, 3, 40d);
        Assert.Equal(20, reopened.Calculate());
        Verify(sheet, edited: true, sourceToo: false);
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes);
        Verify(final.Sheets[0], edited: true, sourceToo: false);
        Assert.Equal(40d, final.Sheets[0].CellAt(2, 3).GetValue<double>());
        Assert.Empty(final.ValidateOpenXml());

        void Verify(ExcelSheet destination, bool edited, bool sourceToo) {
            foreach (JsonElement expected in evidence.GetProperty("cells").EnumerateArray()) {
                int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
                Assert.Equal(expected.GetProperty("formula").GetString(), destination.GetFormulaText(row, column));
                if (sourceToo) {
                    var cell = table.GetCell(row, column)!;
                    Assert.True(cell.FormulaIsComplete && cell.CachedValueIsComplete);
                    Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), Assert.IsType<double>(cell.Value));
                }
                string reference = edited ? "excelEditedXlsx" : "excelRecalculatedXlsx";
                double value = double.Parse(expected.GetProperty(reference).GetProperty("value").GetString()!, CultureInfo.InvariantCulture);
                Assert.Equal(value, destination.CellAt(row, column).GetValue<double>(), precision: 12);
            }
        }
    }
}
