using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyFormulaGroupEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Editing_one_shared_member_materializes_the_remaining_native_relative_formulas(bool encrypted) {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Data");
        for (int row = 1; row <= 3; row++) {
            sheet.CellValue(row, 1, row * 10D); sheet.CellValue(row, 2, row * 10D + 5);
            sheet.CellFormula(row, 2, "A" + row + "+5");
            var formula = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B" + row).CellFormula!;
            formula.FormulaType = CellFormulaValues.Shared; formula.SharedIndex = 0;
            if (row == 1) formula.Reference = "B1:B3"; else formula.Text = string.Empty;
        }
        byte[] bytes = seed.ToBytes(ExcelFileFormat.Xls);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out _));
        Assert.Contains(BiffRecordReader.ReadRecords(compound!.Streams["Workbook"], new()), record => record.Type == 0x04bc);
        using var document = ExcelDocument.Load(new MemoryStream(bytes));
        document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B1").CellFormula = null;
        document.Sheets[0].CellValue(1, 2, 100D); // Remove the native shared-group owner.
        document.Sheets[0].CellBold(3, 2);
        using var output = new MemoryStream();
        if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
        else document.Save(output, ExcelFileFormat.Xls);
        var native = LegacyXlsWorkbook.Load(output.ToArray(), new() { Password = encrypted ? "test-password" : null });
        var remaining = native.Worksheets[0].Cells.Where(cell => cell.Column == 2).OrderBy(cell => cell.Row).ToArray();
        Assert.Equal(100D, remaining[0].Value);
        Assert.Equal("A2+5", remaining[1].FormulaText); Assert.Equal("A3+5", remaining[2].FormulaText);
        var records = BiffRecordReader.ReadRecords(native.SourceCompoundFile!.Streams["Workbook"], new());
        Assert.DoesNotContain(records, record => record.Type == 0x04bc);
        Assert.All(records.Where(record => record.Type == 0x0006), record => Assert.NotEqual(0x01, record.Payload[22]));
        using var loaded = encrypted ? ExcelDocument.LoadEncrypted(output, "test-password") : ExcelDocument.Load(output);
        Assert.True(loaded.Sheets[0].CellAt(3, 2).GetStyle().Bold);
        loaded.Sheets[0].RecalculateSupportedFormulas();
        Assert.Equal(25D, loaded.Sheets[0].CellAt(2, 2).GetValue<double>());
        Assert.Equal(35D, loaded.Sheets[0].CellAt(3, 2).GetValue<double>());
    }

    [Fact]
    public void Clearing_array_caches_keeps_every_native_member_and_the_array_definition() {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Data");
        sheet.CellValue(1, 1, 2D); sheet.CellValue(2, 1, 3D);
        sheet.SetArrayFormula("B1:B2", "SUM(A1:A2)"); sheet.CellValue(1, 2, 5D); sheet.CellValue(2, 2, 5D);
        byte[] original = seed.ToBytes(ExcelFileFormat.Xls);
        foreach (bool reload in new[] { false, true }) {
            using var imported = ExcelDocument.Load(new MemoryStream(original));
            var document = reload ? imported : seed;
            document.Sheets[0].CellBold(2, 2);
            document.Sheets[0].CellValue(4, 1, "Unrelated edit");
            byte[] bytes = document.ToBytes(ExcelFileFormat.Xls, new() { ClearCachedFormulaResultsBeforeSave = true });
            var native = LegacyXlsWorkbook.Load(bytes);
            var array = Assert.Single(native.Worksheets[0].ArrayFormulaRecords);
            Assert.Equal("B1:B2", array.Range); Assert.Equal(2, array.MatchedFormulaCellCount);
            Assert.All(native.Worksheets[0].Cells.Where(cell => cell.Column == 2), cell => Assert.Equal("SUM(A1:A2)", cell.FormulaText));
            using var encrypted = new MemoryStream(); document.SaveEncrypted(encrypted, "test-password", ExcelFileFormat.Xls);
            using var encryptedReadback = ExcelDocument.LoadEncrypted(encrypted, "test-password");
            Assert.Equal(CellFormulaValues.Array, encryptedReadback.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B1").CellFormula?.FormulaType?.Value);
            using var loaded = ExcelDocument.Load(new MemoryStream(bytes));
            Assert.All(loaded.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Where(cell => cell.CellReference?.Value?.StartsWith("B", StringComparison.Ordinal) == true),
                cell => Assert.True(string.IsNullOrEmpty(cell.CellValue?.Text)));
        }
    }
}
