using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelNativeEditingReviewContracts {
    private static string Fixture => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xlsb");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Newly_authored_array_materializes_preexisting_empty_members(bool encrypted) {
        using var seed = ExcelDocument.Create(); var sheet = seed.AddSheet("Data");
        sheet.CellValue(1, 1, 2D); sheet.CellValue(2, 1, 3D); sheet.CellBold(1, 2); sheet.CellBold(2, 2);
        using var document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls })));
        document.Sheets[0].SetArrayFormula("B1:B2", "SUM(A1:A2)");
        using var output = new MemoryStream();
        if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xls);
        else document.Save(output, ExcelFileFormat.Xls);
        var native = LegacyXlsWorkbook.Load(output.ToArray(), new() { Password = encrypted ? "test-password" : null });
        var array = Assert.Single(native.Worksheets[0].ArrayFormulaRecords);
        Assert.Equal("B1:B2", array.Range); Assert.Equal(2, array.MatchedFormulaCellCount);
        Assert.All(native.Worksheets[0].Cells.Where(cell => cell.Column == 2), cell => Assert.True(cell.IsFormula));
        document.Sheets[0].CellValue(3, 1, "second save");
        var second = LegacyXlsWorkbook.Load(document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls }));
        Assert.Equal(2, Assert.Single(second.Worksheets[0].ArrayFormulaRecords).MatchedFormulaCellCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Chartsheet_series_follow_worksheet_rename_through_native_save(bool encrypted) {
        using var fixture = ExcelDocument.Load(Fixture);
        string originalName = fixture.Sheets[0].Name;
        byte[] source = ModifyPackage(File.ReadAllBytes(Fixture), package => {
            var entry = package.GetEntry("xl/charts/chart1.xml")!;
            XDocument chart; using (var input = entry.Open()) chart = XDocument.Load(input);
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var series = chart.Descendants(c + "ser").First();
            series.Element(c + "val")!.ReplaceNodes(new XElement(c + "numRef", new XElement(c + "f", "'" + originalName.Replace("'", "''") + "'!$A$1:$A$3")));
            entry.Delete(); using var output = package.CreateEntry("xl/charts/chart1.xml").Open(); chart.Save(output);
        });
        using var document = ExcelDocument.Load(new MemoryStream(source));
        document.Sheets[0].Name = "Renamed data";
        using var output = new MemoryStream();
        if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xlsb);
        else document.Save(output, ExcelFileFormat.Xlsb);
        using var loaded = encrypted ? ExcelDocument.LoadEncrypted(output, "test-password") : ExcelDocument.Load(output);
        var formulas = loaded.WorkbookPartRoot.ChartsheetParts.SelectMany(part => part.DrawingsPart!.ChartParts)
            .SelectMany(part => part.ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()).Select(formula => formula.Text).ToArray();
        Assert.Contains("'Renamed data'!$A$1:$A$3", formulas);
        Assert.DoesNotContain("'" + originalName.Replace("'", "''") + "'!$A$1:$A$3", formulas);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Signed_VBA_host_changes_fail_before_destination_changes(bool encrypted) {
        byte[] source = ModifyPackage(File.ReadAllBytes(Fixture), package => {
            const string ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            using (var writer = new StreamWriter(package.CreateEntry("xl/_rels/vbaProject.bin.rels").Open()))
                writer.Write("<Relationships xmlns=\"" + ns + "\"><Relationship Id=\"sig\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature\" Target=\"vbaProjectSignature.bin\"/></Relationships>");
            using (var output = package.CreateEntry("xl/vbaProjectSignature.bin").Open()) output.WriteByte(1); // Inert signature carrier; no private key.
            var entry = package.GetEntry("[Content_Types].xml")!;
            XDocument types; using (var input = entry.Open()) types = XDocument.Load(input);
            types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/xl/vbaProjectSignature.bin"),
                new XAttribute("ContentType", "application/vnd.ms-office.vbaProjectSignature")));
            entry.Delete(); using var target = package.CreateEntry("[Content_Types].xml").Open(); types.Save(target);
        });
        using var document = ExcelDocument.Load(new MemoryStream(source));
        document.AddSheet("Added host").CellValue(1, 1, "host");
        byte[] sentinel = { 1, 2, 3, 4 }; using var output = new MemoryStream(); output.Write(sentinel, 0, sentinel.Length);
        var error = Assert.Throws<NotSupportedException>(() => {
            if (encrypted) document.SaveEncrypted(output, "test-password", ExcelFileFormat.Xlsb);
            else document.Save(output, ExcelFileFormat.Xlsb);
        });
        Assert.Contains("sign", error.Message, StringComparison.OrdinalIgnoreCase); Assert.Equal(sentinel, output.ToArray());
    }

    [Fact]
    public void Name_projection_does_not_rebind_cross_scope_or_shadowed_global_tokens() {
        using var seed = ExcelDocument.Create(); var first = seed.AddSheet("First"); var second = seed.AddSheet("Second");
        first.CellValue(1, 1, 10D); second.CellValue(1, 1, 20D);
        seed.SetNamedRange("Value", "'First'!A1", save: false);
        seed.SetNamedRange("Value", "A1", first, save: false);
        seed.SetNamedRange("Value", "A1", second, save: false);
        var model = OfficeIMO.Excel.Xlsb.XlsbWorkbookReader.Load(seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        for (int index = 0; index < model.DefinedNames.Count; index++) {
            var name = model.DefinedNames[index];
            byte[] token = new byte[] { 0x43 }.Concat(BitConverter.GetBytes((uint)index + 1)).ToArray();
            bool decoded = OfficeIMO.Excel.Xlsb.Biff12.XlsbFormulaTextReader.TryRead(token, out string? text, model, 1);
            Assert.Equal(name.LocalSheetIndex == 1, decoded);
            if (decoded) Assert.Equal("Value", text);
        }
    }

    private static byte[] ModifyPackage(byte[] source, Action<ZipArchive> modify) {
        using var stream = new MemoryStream(); stream.Write(source, 0, source.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) modify(archive);
        return stream.ToArray();
    }
}
