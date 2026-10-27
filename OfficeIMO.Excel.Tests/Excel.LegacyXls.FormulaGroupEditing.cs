using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyFormulaGroupEditingTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void Editing_one_shared_member_materializes_the_remaining_native_relative_formulas(bool encrypted, bool crossSheet, bool area) {
        using var seed = ExcelDocument.Create(); var data = seed.AddWorksheet("Data");
        var sheet = crossSheet ? seed.AddWorksheet("Formulas") : data;
        string prefix = crossSheet ? "'Data'!" : string.Empty;
        data.CellValue(4, 1, 40D);
        for (int row = 1; row <= 3; row++) {
            data.CellValue(row, 1, row * 10D); sheet.CellValue(row, 2, row * 10D + 5);
            sheet.CellFormula(row, 2, area ? "SUM(A" + row + ":A" + (row + 1) + ")+5" : "A" + row + "+5");
            var formula = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B" + row).CellFormula!;
            formula.FormulaType = CellFormulaValues.Shared; formula.SharedIndex = 0;
            if (row == 1) formula.Reference = "B1:B3"; else formula.Text = string.Empty;
        }
        byte[] bytes = seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls });
        if (crossSheet) bytes = MakePreviousWriter3dGroup(bytes);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out _));
        Assert.Contains(BiffRecordReader.ReadRecords(compound!.Streams["Workbook"], new()), record => record.Type == 0x04bc);
        using var document = ExcelDocument.Load(new MemoryStream(bytes));
        document.Sheets[crossSheet ? 1 : 0].WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B1").CellFormula = null;
        document.Sheets[crossSheet ? 1 : 0].CellValue(1, 2, 100D); // Remove the native shared-group owner.
        document.Sheets[crossSheet ? 1 : 0].CellBold(3, 2);
        using var output = new MemoryStream();
        if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
        else document.Save(output, ExcelFileFormat.Xls);
        var native = LegacyXlsWorkbook.Load(output.ToArray(), new() { Password = encrypted ? "test-password" : null });
        var remaining = native.Worksheets[crossSheet ? 1 : 0].Cells.Where(cell => cell.Column == 2).OrderBy(cell => cell.Row).ToArray();
        Assert.Equal(100D, remaining[0].Value);
        Assert.Equal(area ? "SUM(" + prefix + "A2:A3)+5" : prefix + "A2+5", remaining[1].FormulaText);
        Assert.Equal(area ? "SUM(" + prefix + "A3:A4)+5" : prefix + "A3+5", remaining[2].FormulaText);
        var records = BiffRecordReader.ReadRecords(native.SourceCompoundFile!.Streams["Workbook"], new());
        Assert.DoesNotContain(records, record => record.Type == 0x04bc);
        Assert.All(records.Where(record => record.Type == 0x0006), record => Assert.NotEqual(0x01, record.Payload[22]));
        using var loaded = encrypted ? ExcelDocument.LoadEncrypted(output, "test-password") : ExcelDocument.Load(output);
        Assert.True(loaded.Sheets[crossSheet ? 1 : 0].CellAt(3, 2).GetStyle().Bold);
        loaded.Sheets[crossSheet ? 1 : 0].RecalculateSupportedFormulas();
        Assert.Equal(area ? 55D : 25D, loaded.Sheets[crossSheet ? 1 : 0].CellAt(2, 2).GetValue<double>());
        Assert.Equal(area ? 75D : 35D, loaded.Sheets[crossSheet ? 1 : 0].CellAt(3, 2).GetValue<double>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Adding_self_references_reuses_the_native_supporting_workbook(bool encrypted) {
        using var seed = ExcelDocument.Create();
        var data = seed.AddWorksheet("Data"); data.CellValue(1, 1, 10D);
        var calculations = seed.AddWorksheet("Calculations"); calculations.CellFormula(1, 1, "Data!A1+1");
        byte[] original = seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls });
        using var document = ExcelDocument.Load(new MemoryStream(original));
        for (int iteration = 0; iteration < 3; iteration++) {
            var added = document.AddWorksheet("Added" + iteration); added.CellValue(1, 1, 20D + iteration);
            document.Sheets[1].CellFormula(iteration + 2, 1, "'" + added.Name + "'!A1+1");
            using var output = new MemoryStream();
            if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
            else document.Save(output, ExcelFileFormat.Xls);
            var native = LegacyXlsWorkbook.Load(output.ToArray(), new() { Password = encrypted ? "test-password" : null });
            var records = BiffRecordReader.ReadRecords(native.SourceCompoundFile!.Streams["Workbook"], new());
            Assert.Single(records, record => record.Type == 0x01ae && record.Payload.Length == 4 && record.Payload[2] == 1 && record.Payload[3] == 4);
            Assert.Equal("'" + added.Name + "'!A1+1", native.Worksheets[1].Cells.Single(cell => cell.Row == iteration + 2 && cell.Column == 1).FormulaText);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Shared_RTD_functions_are_materialized_without_invoking_the_server(bool encrypted, bool nested) {
        const string rtd = "RTD(\"ProgId\",\"Server\",\"Topic\")";
        string expression = nested ? "IF(A1>0," + rtd + ",\"offline\")" : rtd;
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Data");
        for (int row = 1; row <= 2; row++) {
            sheet.CellValue(row, 1, (double)row); sheet.CellValue(row, 2, "snapshot");
            sheet.CellFormula(row, 2, expression);
            var formula = sheet.WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B" + row).CellFormula!;
            formula.FormulaType = CellFormulaValues.Shared; formula.SharedIndex = 0;
            if (row == 1) formula.Reference = "B1:B2"; else formula.Text = string.Empty;
        }
        using var output = new MemoryStream();
        if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
        else document.Save(output, ExcelFileFormat.Xls);
        var native = LegacyXlsWorkbook.Load(output.ToArray(), new() { Password = encrypted ? "test-password" : null });
        Assert.DoesNotContain(BiffRecordReader.ReadRecords(native.SourceCompoundFile!.Streams["Workbook"], new()), record => record.Type == 0x04bc);
        var formulas = native.Worksheets[0].Cells.Where(cell => cell.Column == 2).OrderBy(cell => cell.Row).ToArray();
        Assert.Equal(expression, formulas[0].FormulaText);
        Assert.Equal(expression.Replace("A1", "A2"), formulas[1].FormulaText);
        Assert.All(formulas, cell => Assert.Equal("snapshot", cell.Value));
    }

    [Fact]
    public void Clearing_array_caches_keeps_every_native_member_and_the_array_definition() {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddWorksheet("Data");
        sheet.CellValue(1, 1, 2D); sheet.CellValue(2, 1, 3D);
        sheet.SetArrayFormula("B1:B2", "SUM(A1:A2)"); sheet.CellValue(1, 2, 5D); sheet.CellValue(2, 2, 5D);
        byte[] original = seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls });
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
    // Older OfficeIMO writers placed relative 3D operands in ShrFmla. Keep
    // editing those files safe while the current writer emits ordinary formulas.
    private static byte[] MakePreviousWriter3dGroup(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out _));
        var records = BiffRecordReader.ReadRecords(compound!.Streams["Workbook"], new());
        using var stream = new MemoryStream();
        foreach (var record in records) {
            byte[] payload = record.Payload;
            if (record.Type == 0x04bc) {
                Assert.Contains(payload[10], new byte[] { 0x2c, 0x4c, 0x6c, 0x2d, 0x4d, 0x6d });
                var rewritten = new byte[payload.Length + 2];
                Buffer.BlockCopy(payload, 0, rewritten, 0, 10);
                ushort count = (ushort)(BiffRecordReader.ReadUInt16(payload, 8) + 2);
                rewritten[8] = (byte)count; rewritten[9] = (byte)(count >> 8);
                rewritten[10] = (byte)(payload[10] + 14); // ixti 0 refers to Data; preserve relative row/column offsets.
                Buffer.BlockCopy(payload, 11, rewritten, 13, payload.Length - 11);
                payload = rewritten;
            }
            stream.WriteByte((byte)record.Type); stream.WriteByte((byte)(record.Type >> 8));
            stream.WriteByte((byte)payload.Length); stream.WriteByte((byte)(payload.Length >> 8)); stream.Write(payload, 0, payload.Length);
        }
        return OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Workbook"] = stream.ToArray() });
    }

}
