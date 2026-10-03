using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkWholeAxisFormulaCorpusTests {
    [Fact]
    public void Independent_whole_axis_references_preserve_source_coordinates_and_saved_table_body_ranges() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "whole-axis-formulas.json")));
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(Path.Combine(root, "cross-table-formulas.numbers"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        Assert.False(converted.IsVisualFallback);
        IWorkDiagnostic diagnostic = Assert.Single(converted.Report.Diagnostics,
            d => d.Code == "IWORK_NUMBERS_TABLE_BODY_RANGE_APPROXIMATED");
        Assert.Equal(global::OfficeIMO.OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Throws<InvalidOperationException>(() => converted.Report.RequireNoLoss());
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            IWorkTable table = Assert.Single(Assert.Single(converted.Projection.Sheets,
                s => s.Name == expected.GetProperty("sourceSheet").GetString()).Tables,
                t => t.Name == expected.GetProperty("sourceTable").GetString());
            IWorkTableCell cell = table.GetCell(row, 1)!;
            Assert.True(cell.FormulaIsComplete);
            string function = expected.GetProperty("function").GetString()!;
            string sourceSheet = expected.GetProperty("targetSheet").GetString()!.Replace("'", "''");
            string sourceTable = expected.GetProperty("targetTable").GetString()!.Replace("'", "''");
            Assert.Equal("=" + function + "('" + sourceSheet + "'::'" + sourceTable + "'::"
                + expected.GetProperty("sourceCoordinateRange").GetString() + ")", cell.Formula);
            string sourceName = Assert.Single(converted.WorksheetMappings, m => m.SourceTableName == table.Name).DestinationName;
            string target = Assert.Single(converted.WorksheetMappings, m =>
                m.SourceSheetName == expected.GetProperty("targetSheet").GetString()
                && m.SourceTableName == expected.GetProperty("targetTable").GetString()).DestinationName;
            ExcelSheet sheet = Assert.Single(reopened.Sheets, s => s.Name == sourceName);
            Assert.Equal(function + "('" + target.Replace("'", "''") + "'!"
                + expected.GetProperty("destinationBodyRange").GetString() + ")", sheet.GetFormulaText(row, 1));
            Assert.Equal(expected.GetProperty("cachedValue").GetDouble(), sheet.CellAt(row, 1).GetValue<double>());
            Assert.Equal(expected.GetProperty("cachedValue").GetDouble(), expected.GetProperty("computedCurrentBodyValue").GetDouble());
        }
    }
}
