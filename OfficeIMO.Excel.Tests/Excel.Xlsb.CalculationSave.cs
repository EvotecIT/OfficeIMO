using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelXlsbCalculationSaveTests {
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
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xlsb");
        using var document = ExcelDocument.Load(new MemoryStream(File.ReadAllBytes(path)));
        var sheet = document.Sheets[0]; sheet.CellFormula(200, 4, "1+2");
        sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "D200").CellValue = new CellValue("99");
        var options = new ExcelSaveOptions {
            EvaluateFormulasBeforeSave = mode == "evaluate", ClearCachedFormulaResultsBeforeSave = mode == "clear",
            MarkFormulasDirtyBeforeSave = mode == "dirty", ForceFullCalculationOnOpen = mode == "full"
        };
        using var output = new MemoryStream();
        if (encryptedOutput) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xlsb, options);
        else document.Save(output, ExcelFileFormat.Xlsb, options);
        using var loaded = encryptedOutput ? ExcelDocument.LoadEncrypted(output, "test-password") : ExcelDocument.Load(output);
        Assert.Equal(4, loaded.WorkbookPartRoot.ChartsheetParts.Count());
        Assert.Equal(document.ExtractMacros(), loaded.ExtractMacros());
        var cell = loaded.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == "D200");
        Assert.Equal("1+2", cell.CellFormula?.Text);
        if (mode == "clear") Assert.Equal(string.Empty, loaded.Sheets[0].GetCellText(cell));
        else Assert.Equal(mode == "evaluate" ? "3" : "99", cell.CellValue?.Text);
        if (mode != "evaluate") {
            Assert.True(loaded.WorkbookRoot.CalculationProperties?.FullCalculationOnLoad?.Value);
            var native = XlsbWorkbookReader.Load(loaded.ToBytes(ExcelFileFormat.Xlsb)).Worksheets[0].Cells.Single(item => item.Row == 200 && item.Column == 4);
            Assert.Equal(2, native.FormulaPayloadBytes![0] & 3);
        }
    }

    [Fact]
    public void Fresh_workbooks_and_document_level_calculation_settings_apply_to_native_saves() {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Data");
        sheet.CellFormula(1, 1, "1+2"); sheet.RecalculateSupportedFormulas();
        byte[] cleared = seed.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { ClearCachedFormulaResultsBeforeSave = true });
        using var blank = ExcelDocument.Load(new MemoryStream(cleared));
        Assert.Equal(string.Empty, blank.Sheets[0].GetCellText(blank.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single()));
        Assert.True(blank.WorkbookRoot.CalculationProperties?.FullCalculationOnLoad?.Value);
        blank.Calculation.EvaluateFormulasBeforeSave = true;
        using var calculated = ExcelDocument.Load(new MemoryStream(blank.ToBytes(ExcelFileFormat.Xlsb)));
        Assert.Equal(3D, calculated.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Clearing_opaque_native_formula_cache_keeps_tokens_on_a_subsequent_edit() {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Data");
        sheet.CellFormula(1, 1, "1+2"); sheet.RecalculateSupportedFormulas();
        byte[] source = seed.ToBytes(ExcelFileFormat.Xlsb);
        byte[] refError = { 0x4A, 0, 0, 0, 0, 0, 0 }; // Valid BIFF12 PtgRefErrV, deliberately outside the text projection subset.
        using var edited = new MemoryStream(); edited.Write(source, 0, source.Length);
        using (var archive = new ZipArchive(edited, ZipArchiveMode.Update, leaveOpen: true)) {
            var entry = archive.GetEntry("xl/worksheets/sheet1.bin")!;
            IReadOnlyList<XlsbRecord> records;
            using (var input = entry.Open()) records = XlsbRecordReader.ReadAll(input);
            using var output = new MemoryStream();
            foreach (var record in records) {
                byte[] payload = record.Data;
                if (record.Type == 9) {
                    payload = new byte[22 + refError.Length + 4]; Buffer.BlockCopy(record.Data, 0, payload, 0, 18);
                    BitConverter.GetBytes(refError.Length).CopyTo(payload, 18); refError.CopyTo(payload, 22);
                }
                XlsbRecordWriter.Write(output, record.Type, payload);
            }
            entry.Delete(); using var target = archive.CreateEntry("xl/worksheets/sheet1.bin").Open(); byte[] bytes = output.ToArray(); target.Write(bytes, 0, bytes.Length);
        }
        using var document = ExcelDocument.Load(new MemoryStream(edited.ToArray()));
        Assert.Null(document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single().CellFormula);
        using var cleared = new MemoryStream();
        document.Save(cleared, ExcelFileFormat.Xlsb, new ExcelSaveOptions { ClearCachedFormulaResultsBeforeSave = true });
        var first = XlsbWorkbookReader.Load(cleared.ToArray()).Worksheets[0].Cells.Single();
        Assert.Equal(refError, first.FormulaBytes); Assert.Equal(string.Empty, first.Value);
        document.Sheets[0].CellValue(2, 1, "later edit");
        var second = XlsbWorkbookReader.Load(document.ToBytes(ExcelFileFormat.Xlsb)).Worksheets[0].Cells.Single(cell => cell.Row == 1);
        Assert.Equal(refError, second.FormulaBytes); Assert.Equal(string.Empty, second.Value);
    }
}
