using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkSingleCellFormulaCorpusTests {
    [Fact]
    public void Independent_single_cell_source_coordinates_and_typed_caches_are_qualified_without_bypassing_date_safety() {
        using var manifest = ReadManifest("single-cell-formulas", out string path);
        IWorkSourceDocument source = IWorkSourceDocument.Open(path);
        IWorkNumbersProjection projection = source.ReadNumbers();
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray())
            AssertSourceCell(projection, expected);
        using var result = source.ToExcelDocumentResult(new IWorkConversionOptions {
            AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED"
            && diagnostic.Message.Contains("date outside", StringComparison.Ordinal));
    }

    [Fact]
    public void Independent_endpoint_ranges_and_single_cell_offset_preserve_saved_formulas_caches_and_edited_calculation() {
        using var manifest = ReadManifest("endpoint-formulas", out string path);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        Assert.False(result.IsVisualFallback, string.Join("\n", result.Report.Diagnostics.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message)));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            IWorkTable table = AssertSourceCell(result.Projection, expected);
            int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
            IWorkFormulaCellStatus status = Assert.Single(result.Report.FormulaCells, status =>
                status.TableIdentity?.RecordIdentifier == table.SourceIdentity!.RecordIdentifier && status.Row == row && status.Column == column);
            Assert.True(status.ExpressionIsComplete); Assert.Equal(IWorkFormulaCacheStatus.Complete, status.CacheStatus);
            string sourceName = Name(expected.GetProperty("sourceSheet").GetString()!, expected.GetProperty("sourceTable").GetString()!);
            string targetName = Name(expected.GetProperty("targetSheet").GetString()!, expected.GetProperty("targetTable").GetString()!);
            ExcelSheet sheet = reopened.Sheets.Single(sheet => sheet.Name == sourceName);
            Assert.Equal(expected.GetProperty("sourceFormula").GetString()!.Replace("Data::", "'" + targetName.Replace("'", "''") + "'!"),
                sheet.GetFormulaText(row, column));
            Assert.Equal(expected.GetProperty("cachedValue").GetDouble(), sheet.CellAt(row, column).GetValue<double>());
        }
        ExcelSheet formulas = reopened.Sheets.Single(sheet => sheet.Name == Name("Formulas", "Tests"));
        ExcelSheet data = reopened.Sheets.Single(sheet => sheet.Name == Name("Formulas", "Data"));
        data.CellValue(3, 3, 80d);
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal(80d, formulas.CellAt(79, 2).GetValue<double>());
        Assert.Equal(80d, formulas.CellAt(102, 2).GetValue<double>());
        Assert.Equal(2d, formulas.CellAt(29, 2).GetValue<double>());

        string Name(string sourceSheet, string sourceTable) => result.WorksheetMappings.Single(mapping =>
            mapping.SourceSheetName == sourceSheet && mapping.SourceTableName == sourceTable).DestinationName;
    }

    private static JsonDocument ReadManifest(string name, out string path) {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser");
        path = Path.Combine(root, name + ".numbers");
        JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, name + ".json")));
        Assert.Equal(manifest.RootElement.GetProperty("sourceSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        return manifest;
    }

    private static IWorkTable AssertSourceCell(IWorkNumbersProjection projection, JsonElement expected) {
        IWorkTable table = projection.Sheets.Single(sheet => sheet.Name == expected.GetProperty("sourceSheet").GetString())
            .Tables.Single(table => table.Name == expected.GetProperty("sourceTable").GetString());
        IWorkTableCell cell = table.GetCell(expected.GetProperty("row").GetInt32(), expected.GetProperty("column").GetInt32())!;
        Assert.True(cell.FormulaIsComplete);
        string qualifier = "'" + expected.GetProperty("targetSheet").GetString()!.Replace("'", "''") + "'::'"
            + expected.GetProperty("targetTable").GetString()!.Replace("'", "''") + "'::";
        Assert.Equal("=" + expected.GetProperty("sourceFormula").GetString()!.Replace("Data::", qualifier), cell.Formula);
        object value = expected.GetProperty("cachedValue").ValueKind == JsonValueKind.String
            ? expected.GetProperty("cachedValue").GetString()! : expected.GetProperty("cachedValue").GetDouble();
        Assert.Equal(value, cell.Value); Assert.True(cell.CachedValueIsComplete);
        object computed = expected.GetProperty("computedCurrentValue").ValueKind == JsonValueKind.String
            ? expected.GetProperty("computedCurrentValue").GetString()! : expected.GetProperty("computedCurrentValue").GetDouble();
        Assert.Equal(value, computed);
        return table;
    }
}
