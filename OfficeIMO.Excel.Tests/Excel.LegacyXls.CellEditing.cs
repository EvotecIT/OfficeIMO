using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyCellEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Cell_add_delete_type_formula_and_style_changes_keep_native_chart_and_VBA_carriers(bool encryptedSource) {
        string fixture = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");
        byte[] originalBytes = File.ReadAllBytes(fixture);
        using var original = new MemoryStream(originalBytes);
        using var source = ExcelDocument.Load(original);
        using var encrypted = new MemoryStream();
        if (encryptedSource) source.SaveEncrypted(encrypted, "test-password", ExcelFileFormat.Xls);
        using var document = encryptedSource ? ExcelDocument.LoadEncrypted(encrypted, "test-password") : ExcelDocument.Load(new MemoryStream(originalBytes));
        var sheet = document.Sheets[0];
        string deletedAddress = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().First(cell => cell.CellReference?.Value != "A1" && cell.CellFormula == null).CellReference!.Value!;
        sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == deletedAddress).Remove();
        sheet.CellValue(1, 1, 3.141592653589793D);
        sheet.CellValue(200, 3, true);
        sheet.CellValue(201, 3, "Added Unicode Łódź value");
        sheet.CellFormula(200, 4, "1+2");
        sheet.RecalculateSupportedFormulas();
        sheet.CellBold(201, 3);
        sheet.CellValue(202, 3, "removed cell");
        sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "C202").Remove();
        byte[] saved = document.ToBytes(ExcelFileFormat.Xls);
        using var reloaded = ExcelDocument.Load(new MemoryStream(saved));
        Assert.Equal(3.141592653589793D, reloaded.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.True(reloaded.Sheets[0].CellAt(200, 3).GetValue<bool>());
        Assert.Equal("Added Unicode Łódź value", reloaded.Sheets[0].CellAt(201, 3).GetValue<string>());
        Assert.True(reloaded.Sheets[0].CellAt(201, 3).GetStyle().Bold);
        Assert.Equal("1+2", reloaded.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "D200").CellFormula?.Text);
        Assert.DoesNotContain(reloaded.Sheets[0].WorksheetPart.Worksheet.Descendants<Cell>(), cell => cell.CellReference?.Value == "C202" || cell.CellReference?.Value == deletedAddress);
        Assert.Equal(4, reloaded.LegacyXlsChartSheets.Count);
        Assert.True(OfficeCompoundFileReader.TryRead(originalBytes, out OfficeCompoundFile? before, out string? error), error);
        Assert.True(OfficeCompoundFileReader.TryRead(saved, out OfficeCompoundFile? after, out error), error);
        foreach (var stream in before!.Streams.Where(item => item.Key != "Workbook")) Assert.Equal(stream.Value, after!.Streams[stream.Key]);
        var oldRecords = BiffRecordReader.ReadRecords(before.Streams["Workbook"], new());
        var newRecords = BiffRecordReader.ReadRecords(after!.Streams["Workbook"], new());
        foreach (ushort type in new ushort[] { 0x0031, 0x00e0 })
            Assert.Equal(oldRecords.Where(record => record.Type == type).Select(record => record.Payload),
                newRecords.Where(record => record.Type == type).Take(oldRecords.Count(record => record.Type == type)).Select(record => record.Payload));
    }
}
