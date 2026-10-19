using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyStructureEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Worksheet_lifecycle_keeps_chart_carriers_and_reconciles_VBA_document_modules(bool encryptedSource) {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "excel-populated-macro.xls");
        byte[] original = File.ReadAllBytes(path);
        using var seed = ExcelDocument.Load(new MemoryStream(original));
        using var encrypted = new MemoryStream();
        if (encryptedSource) seed.SaveEncrypted(encrypted, "test-password", ExcelFileFormat.Xls);
        using var document = encryptedSource ? ExcelDocument.LoadEncrypted(encrypted, "test-password") : ExcelDocument.Load(new MemoryStream(original));
        string removedName = document.Sheets[1].Name;
        var model = LegacyXlsWorkbook.Load(new MemoryStream(original));
        string? removedCodeName = model.Worksheets.Single(sheet => sheet.Name == removedName).CodeName;
        var retained = document.Sheets[0]; retained.Name = "Sales Łódź"; retained.CellValue(1, 1, 42D);
        document.ReorderWorksheet(retained, document.WorkbookRoot.Sheets!.Elements<Sheet>().Count() - 1);
        document.RemoveWorksheet(removedName);
        var added = document.AddWorksheet("Added data");
        added.CellValue(1, 1, "New worksheet host");
        added.CellFormula(2, 1, "'Sales Łódź'!A1+1"); added.RecalculateSupportedFormulas();
        byte[] bytes = document.ToBytes(ExcelFileFormat.Xls);
        using var loaded = ExcelDocument.Load(new MemoryStream(bytes));
        Assert.Equal(document.Sheets.Select(sheet => sheet.Name), loaded.Sheets.Select(sheet => sheet.Name));
        Assert.Equal(4, loaded.LegacyXlsChartSheets.Count);
        Assert.Equal(42D, loaded.GetSheet("Sales Łódź").CellAt(1, 1).GetValue<double>());
        Assert.Equal(43D, loaded.GetSheet("Added data").CellAt(2, 1).GetValue<double>());
        var result = LegacyXlsWorkbook.Load(new MemoryStream(bytes));
        string? addedCodeName = result.Worksheets.Single(sheet => sheet.Name == "Added data").CodeName;
        Assert.False(string.IsNullOrEmpty(addedCodeName));
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? before, out string? error), error);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? after, out error), error);
        var beforeProject = ReadProject(before!); var afterProject = ReadProject(after!);
        Assert.DoesNotContain(afterProject.Modules, module => module.Name == removedCodeName);
        Assert.Contains(afterProject.Modules, module => module.Name == addedCodeName && module.Kind == OfficeVbaModuleKind.Document
            && module.Source.Contains("00020820-0000-0000-C000-000000000046"));
        foreach (var module in beforeProject.Modules.Where(module => module.Name != removedCodeName))
            Assert.Equal(module.Source, afterProject.GetModule(module.Name).Source);
    }

    [Fact]
    public void Reordering_keeps_local_names_attached_and_removing_a_named_sheet_does_not_retarget_name_tokens() {
        using var seed = ExcelDocument.Create();
        var first = seed.AddWorksheet("First"); first.CellValue(1, 1, 10);
        var second = seed.AddWorksheet("Second"); second.CellValue(1, 1, 20);
        var third = seed.AddWorksheet("Third"); third.CellValue(1, 1, 30);
        seed.SetNamedRange("FirstLocal", "A1", first, save: false);
        seed.SetNamedRange("SecondLocal", "A1", second, save: false);
        seed.SetNamedRange("GlobalThird", "'Third'!A1", save: false);
        third.CellFormula(2, 1, "GlobalThird+1"); third.RecalculateSupportedFormulas();
        using var document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(ExcelFileFormat.Xls)));
        document.ReorderWorksheet("Second", 0);
        Assert.Equal(0U, document.WorkbookRoot.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "SecondLocal").LocalSheetId?.Value);
        document.RemoveWorksheet("First");
        byte[] bytes = document.ToBytes(ExcelFileFormat.Xls);
        using var loaded = ExcelDocument.Load(new MemoryStream(bytes));
        Assert.Equal(new[] { "Second", "Third" }, loaded.Sheets.Select(sheet => sheet.Name));
        Assert.DoesNotContain(loaded.WorkbookRoot.DefinedNames!.Elements<DefinedName>(), name => name.Name == "FirstLocal");
        Assert.Equal(0U, loaded.WorkbookRoot.DefinedNames.Elements<DefinedName>().Single(name => name.Name == "SecondLocal").LocalSheetId?.Value);
        Assert.Equal("GlobalThird+1", loaded.GetSheet("Third").WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2").CellFormula?.Text);
    }

    private static OfficeVbaProject ReadProject(OfficeCompoundFile source) {
        var storage = OfficeCompoundStorage.Extract(source, "_VBA_PROJECT_CUR");
        return OfficeVbaProject.Load(OfficeCompoundFileWriter.Rewrite(storage, new Dictionary<string, byte[]>()));
    }
}
