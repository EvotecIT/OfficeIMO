using OfficeIMO.Excel;
using OfficeIMO.Excel.Legacy;
using System.IO.Compression;
using System.Xml.Linq;

namespace OfficeIMO.LegacyImport.Tests;

public sealed class LaterSpreadsheetContractTests {
    [Theory]
    [InlineData(0, 0x64, "ROWS")]
    [InlineData(0, 0x6b, "PROPER")]
    [InlineData(1, 0x64, "ROWS")]
    [InlineData(1, 0x6b, "PROPER")]
    [InlineData(2, 0x64, "ROWS")]
    [InlineData(2, 0x6b, "PROPER")]
    [InlineData(3, 0x64, "ROWS")]
    [InlineData(3, 0x6b, "PROPER")]
    public void NativeFunctionOpcodesKeepTheirMeaningAcrossSharedDecoderProfiles(int profile, byte opcode, string function) {
        byte[] source;
        if (profile == 0) source = LegacyFixtureFactory.Wk(formulaTokens: new byte[] { 1, 0, 0, 0, 0, opcode, 3 });
        else if (profile == 1) source = Join(Record(0, LotusHeader()), Record(0x28,
            Join(new byte[] { 0, 0, 0, 1 }, BitConverter.GetBytes(1d), new byte[] { 1, 0, 0, 0, 0, 0, opcode, 3 })), Record(1));
        else if (profile == 2) source = Join(Record(0, new byte[] { 1, 0x10 }), Record(0x10,
            Join(new byte[6], BitConverter.GetBytes(1d), new byte[2], new byte[] { 9, 0, 3, 0, 1, opcode, 3, 0, 0, 0, 0, 0, 0 })), Record(1));
        else {
            byte[] tokens = { 1, 0, 0, 0, 0, opcode, 3 };
            source = Join(Record(0xff, new byte[] { 4, 4 }), Record(0x10,
                Join(new byte[6], BitConverter.GetBytes(1d), BitConverter.GetBytes((ushort)tokens.Length), tokens)), Record(1));
        }
        using var result = Import(source);
        Assert.Equal(function + "($A$1)", Assert.Single(result.Cells, cell => cell.Formula != null).Formula);
    }

    [Theory]
    [InlineData("Other__Sheet", "Other__Sheet")]
    [InlineData(" Other ", "Other")]
    [InlineData("Sheet3", "Sheet3")]
    public void FormulaReferencesKeepFinalUniqueWorksheetNamesThroughXlsx(string name, string projected) {
        byte[] source = Join(Record(0, LotusHeader()), Record(0x23, Join(new byte[] { 0xb0, 0x36, 1, 0 }, Encoding.ASCII.GetBytes(name + "\0"))),
            LotusFormula(0, 1), LotusFormula(1, 2), Record(1));
        using var result = Import(source);
        string generated = projected == "Sheet3" ? "Sheet3 (2)" : "Sheet3";
        Assert.Equal(new[] { "Sheet1", projected, generated }, result.Value.Sheets.Select(sheet => sheet.Name));
        Assert.Equal("'" + projected + "'!$A$1", result.Cells[0].Formula);
        Assert.Equal("'" + generated + "'!$A$1", result.Cells[1].Formula);
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using (var loaded = ExcelDocument.Load(output)) Assert.Equal(new[] { "Sheet1", projected, generated }, loaded.Sheets.Select(sheet => sheet.Name));
        output.Position = 0;
        using var package = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true);
        using var sheetXml = package.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Assert.Equal(result.Cells.Select(cell => cell.Formula), XDocument.Load(sheetXml).Descendants(ns + "f").Select(formula => formula.Value));
    }

    [Fact]
    public void LotusBofKeepsDeclaredEmptySheetsWithinTheItemLimit() {
        byte[] header = LotusHeader(); header[10] = 2; header[11] = 1;
        byte[] source = Join(Record(0, header), Record(0x16, Join(new byte[4], Encoding.ASCII.GetBytes("'First\0"))), Record(1));
        using var result = Import(source);
        Assert.Equal(3, result.Value.Sheets.Count);
        Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, result.Value.Sheets.Select(sheet => sheet.Name));
        Assert.Single(result.Cells);
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(source,
            new LegacySpreadsheetImportOptions { RequireStructured = true, Limits = new OfficeLegacyImportLimits { MaxItems = 2 } }));
    }
    private static byte[] LotusFormula(byte column, byte sheet) => Record(0x28,
        Join(new byte[] { 0, 0, 0, column }, BitConverter.GetBytes(1d), new byte[] { 1, 0, 0, 0, sheet, 0, 3 }));
    private static byte[] LotusHeader() { var header = new byte[26]; header[0] = 5; header[1] = 0x10; return header; }
    private static LegacySpreadsheetImportResult Import(byte[] source) => LegacySpreadsheetImporter.Import(source, new LegacySpreadsheetImportOptions { RequireStructured = true });
    private static byte[] Record(ushort type, params byte[] payload) => Join(BitConverter.GetBytes(type), BitConverter.GetBytes((ushort)payload.Length), payload);
    private static byte[] Join(params byte[][] values) => values.SelectMany(value => value).ToArray();
}
