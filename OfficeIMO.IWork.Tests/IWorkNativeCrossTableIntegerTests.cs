using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNativeCrossTableIntegerTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_cross_table_integer_calls_keep_caches_identities_and_saved_recalculation(bool changed) {
        using var manifest = JsonDocument.Parse(File.ReadAllText(Fixture("numbers-integer-cross-table-v14.5-decode.json")));
        var root = manifest.RootElement;
        foreach (var input in root.GetProperty("inputs").EnumerateArray()) {
            byte[] bytes = File.ReadAllBytes(Fixture(input.GetProperty("path").GetString()!));
            Assert.Equal(input.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        string path = Fixture(root.GetProperty(changed ? "changedSource" : "source").GetString()!);
        Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly }));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED");
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
        IWorkTable[] tables = result.Projection.Sheets.SelectMany(sheet => sheet.Tables).ToArray();
        Assert.Equal(3, tables.Length);
        Assert.Equal(root.GetProperty(changed ? "changedSourceTables" : "sourceTables").EnumerateArray().Select(t => t.GetString()),
            tables.Select(table => table.Name));
        Assert.All(result.Projection.Sheets, sheet => Assert.Empty(sheet.TextBoxes));
        Assert.Equal(9, tables.Sum(table => table.Cells.Count(cell => cell.Kind == IWorkCellKind.Formula && cell.FormulaIsComplete)));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(tables.Select(table => table.Name == "Empty" ? "Empty grid" : "Budget - " + table.Name),
            reopened.Sheets.Select(sheet => sheet.Name));
        Verify(changed ? "changedCells" : "cells", true);
        Assert.Equal(9, reopened.Calculate()); Verify(changed ? "changedCells" : "cells", false);
        reopened.Sheets.Single(sheet => sheet.Name == "Budget - Data").CellValue(2, 2, 24d);
        Assert.Equal(9, reopened.Calculate()); Verify("changedCells", false);
        using var finalBytes = new MemoryStream(); reopened.Save(finalBytes); finalBytes.Position = 0;
        using var final = ExcelDocument.Load(finalBytes);
        Assert.Equal(12d, final.Sheets[2].CellAt(1, 1).GetValue<double>());
        Assert.Equal(360d, final.Sheets[2].CellAt(1, 2).GetValue<double>());
        Assert.Equal(792d, final.Sheets[2].CellAt(2, 1).GetValue<double>());
        Assert.Empty(final.ValidateOpenXml());

        void Verify(string field, bool sourceToo) {
            foreach (var expected in root.GetProperty(field).EnumerateArray()) {
                int index = Array.FindIndex(tables, table => table.Name == expected.GetProperty("tableName").GetString()),
                    row = expected.GetProperty("row").GetInt32(),
                    column = expected.GetProperty("column").GetInt32();
                if (sourceToo) {
                    var cell = tables[index].GetCell(row, column)!;
                    Assert.True(cell.FormulaIsComplete); Assert.True(cell.CachedValueIsComplete);
                    Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), cell.Value);
                }
                Assert.Equal(expected.GetProperty("excelFormula").GetString(), reopened.Sheets[index].GetFormulaText(row, column));
                Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), reopened.Sheets[index].CellAt(row, column).GetValue<double>());
            }
        }
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
