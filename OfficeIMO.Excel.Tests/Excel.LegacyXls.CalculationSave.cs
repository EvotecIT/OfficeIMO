using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Core.Internal;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyCalculationSaveTests {
    [Theory]
    [InlineData("evaluate", false)]
    [InlineData("clear", false)]
    [InlineData("dirty", false)]
    [InlineData("full", false)]
    [InlineData("evaluate", true)]
    [InlineData("clear", true)]
    [InlineData("dirty", true)]
    [InlineData("full", true)]
    public void Save_calculation_policy_preserves_native_chart_and_VBA_graphs(string mode, bool encryptedOutput) {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");
        using var document = ExcelDocument.Load(new MemoryStream(File.ReadAllBytes(path)));
        var sheet = document.Sheets[0]; sheet.CellFormula(200, 4, "1+2");
        sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "D200").CellValue = new CellValue("99");
        var options = new ExcelSaveOptions {
            EvaluateFormulasBeforeSave = mode == "evaluate", ClearCachedFormulaResultsBeforeSave = mode == "clear",
            MarkFormulasDirtyBeforeSave = mode == "dirty", ForceFullCalculationOnOpen = mode == "full"
        };
        using var output = new MemoryStream();
        if (encryptedOutput) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls, options);
        else document.Save(output, ExcelFileFormat.Xls, options);
        using var loaded = encryptedOutput ? ExcelDocument.LoadEncrypted(output, "test-password") : ExcelDocument.Load(output);
        Assert.Equal(4, loaded.LegacyXlsChartSheets.Count);
        var nativeWorkbook = LegacyXlsWorkbook.Load(output.ToArray(), new LegacyXlsImportOptions { Password = encryptedOutput ? "test-password" : null });
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(path), out OfficeCompoundFile? before, out _));
        foreach (var stream in before!.Streams.Where(item => item.Key != "Workbook")) Assert.Equal(stream.Value, nativeWorkbook.SourceCompoundFile!.Streams[stream.Key]);
        var cell = loaded.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == "D200");
        Assert.Equal("1+2", cell.CellFormula?.Text);
        if (mode == "clear") Assert.Equal(string.Empty, loaded.Sheets[0].GetCellText(cell));
        else Assert.Equal(mode == "evaluate" ? "3" : "99", cell.CellValue?.Text);
        if (mode != "evaluate") {
            Assert.True(loaded.Sheets[0].WorksheetPart.Worksheet!.GetFirstChild<SheetCalculationProperties>()?.FullCalculationOnLoad?.Value);
            var records = BiffRecordReader.ReadRecords(nativeWorkbook.SourceCompoundFile!.Streams["Workbook"], new());
            var formula = records.First(record => record.Type == 0x0006 && BiffRecordReader.ReadUInt16(record.Payload, 0) == 199 && BiffRecordReader.ReadUInt16(record.Payload, 2) == 3);
            Assert.Equal(1, formula.Payload[14] & 3);
        }
    }

    [Fact]
    public void Fresh_workbooks_and_document_level_calculation_settings_apply_to_native_saves() {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Data");
        sheet.CellFormula(1, 1, "1+2"); sheet.RecalculateSupportedFormulas();
        byte[] cleared = seed.ToBytes(ExcelFileFormat.Xls, new ExcelSaveOptions { ClearCachedFormulaResultsBeforeSave = true });
        using var blank = ExcelDocument.Load(new MemoryStream(cleared));
        Assert.Equal(string.Empty, blank.Sheets[0].GetCellText(blank.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single()));
        Assert.True(blank.Sheets[0].WorksheetPart.Worksheet!.GetFirstChild<SheetCalculationProperties>()?.FullCalculationOnLoad?.Value);
        blank.Calculation.EvaluateFormulasBeforeSave = true;
        using var calculated = ExcelDocument.Load(new MemoryStream(blank.ToBytes(ExcelFileFormat.Xls)));
        Assert.Equal(3D, calculated.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

}
