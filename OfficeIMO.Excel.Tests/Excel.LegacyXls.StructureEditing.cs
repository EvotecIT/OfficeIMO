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
    public void Adding_empty_worksheets_preserves_row_and_column_layout_styles(bool explicitStyle) {
        using var seed = ExcelDocument.Create();
        var original = seed.AddWorksheet("Original");
        original.CellValue(1, 1, 42);
        original.CellAt(1, 1).SetNumberFormat("0.000");
        using var document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls })));
        uint style = explicitStyle
            ? document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>().Single().StyleIndex!.Value : 0U;
        var added = document.AddWorksheet("Layout only");
        added.SetColumnWidth(2, ExcelColumnWidth.Characters(12.5));
        added.SetRowHeight(3, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(21)));
        if (explicitStyle) {
            added.WorksheetPart.Worksheet!.GetFirstChild<Columns>()!.Elements<Column>().Single().Style = style;
            var row = added.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Single();
            row.StyleIndex = style;
            row.CustomFormat = true;
        }
        using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls })));
        var result = loaded.GetSheet("Layout only");
        Assert.Equal(12.5, Assert.Single(result.GetColumnDefinitions()).Width?.CharacterCount);
        Assert.Equal(21, Assert.Single(result.GetRowDefinitions()).Height?.PhysicalLength?.ToPoints());
        var native = LegacyXlsWorkbook.Load(document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls }));
        var layout = native.Worksheets.Single(sheet => sheet.Name == "Layout only");
        Assert.Empty(layout.Cells);
        if (explicitStyle) {
            Assert.Equal("0.000", native.CellFormats[Assert.Single(layout.Columns).StyleIndex].NumberFormatCode);
            Assert.Equal("0.000", native.CellFormats[Assert.Single(layout.Rows).StyleIndex!.Value].NumberFormatCode);
        }
    }

    [Fact]
    public void Native_save_as_keeps_following_saves_on_the_selected_destination() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Xls-Association-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string originalPath = Path.Combine(root, "original.xls"), target = Path.Combine(root, "edited.xls");
        try {
            using var seed = ExcelDocument.Create(); seed.AddWorksheet("Data").CellValue(1, 1, "Original");
            byte[] original = seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls }); File.WriteAllBytes(originalPath, original);
            using (var document = ExcelDocument.Load(originalPath)) {
                document.Sheets[0].CellValue(1, 1, "First edit"); document.Save(target);
                Assert.Equal(target, document.FilePath);
                document.Sheets[0].CellValue(1, 1, "Second edit"); document.Save();
            }
            Assert.Equal(original, File.ReadAllBytes(originalPath));
            using var loaded = ExcelDocument.Load(target);
            Assert.Equal("Second edit", loaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        } finally { Directory.Delete(root, recursive: true); }
    }

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
        byte[] bytes = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls });
        using var loaded = ExcelDocument.Load(new MemoryStream(bytes));
        Assert.Equal(document.Sheets.Select(sheet => sheet.Name), loaded.Sheets.Select(sheet => sheet.Name));
        Assert.Equal(4, loaded.LegacyXlsChartSheets.Count);
        Assert.Equal(42D, loaded.GetSheet("Sales Łódź").CellAt(1, 1).GetValue<double>());
        Assert.Equal(43D, loaded.GetSheet("Added data").CellAt(2, 1).GetValue<double>());
        var result = LegacyXlsWorkbook.Load(new MemoryStream(bytes));
        string? addedCodeName = result.Worksheets.Single(sheet => sheet.Name == "Added data").CodeName;
        foreach (var native in result.Worksheets)
            Assert.Equal(native.CodeName, loaded.GetSheet(native.Name).WorksheetPart.Worksheet!.SheetProperties?.CodeName?.Value);
        if (result.Windows.Count > 0) Assert.Equal((uint)result.Windows[0].ActiveSheetIndex, loaded.WorkbookRoot.BookViews!.Elements<WorkbookView>().First().ActiveTab?.Value);

        Assert.False(string.IsNullOrEmpty(addedCodeName));
        Assert.Equal(addedCodeName, document.GetSheet("Added data").WorksheetPart.Worksheet!.SheetProperties?.CodeName?.Value);
        document.GetSheet("Added data").CellValue(3, 1, "Second native save");
        using var secondSave = ExcelDocument.Load(new MemoryStream(document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls })));
        Assert.Equal("Second native save", secondSave.GetSheet("Added data").CellAt(3, 1).GetValue<string>());
        Assert.Equal(addedCodeName, secondSave.GetSheet("Added data").WorksheetPart.Worksheet!.SheetProperties?.CodeName?.Value);
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
        using var document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls })));
        document.ReorderWorksheet("Second", 0);
        Assert.Equal(0U, document.WorkbookRoot.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "SecondLocal").LocalSheetId?.Value);
        document.RemoveWorksheet("First");
        byte[] bytes = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls });
        using var loaded = ExcelDocument.Load(new MemoryStream(bytes));
        Assert.Equal(new[] { "Second", "Third" }, loaded.Sheets.Select(sheet => sheet.Name));
        Assert.DoesNotContain(loaded.WorkbookRoot.DefinedNames!.Elements<DefinedName>(), name => name.Name == "FirstLocal");
        Assert.Equal(0U, loaded.WorkbookRoot.DefinedNames.Elements<DefinedName>().Single(name => name.Name == "SecondLocal").LocalSheetId?.Value);
        Assert.Equal("GlobalThird+1", loaded.GetSheet("Third").WorksheetPart.Worksheet!.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2").CellFormula?.Text);
    }

    [Theory]
    [InlineData(ExcelFileFormat.Xls)]
    [InlineData(ExcelFileFormat.Xlsb)]
    public void Renaming_a_native_sheet_keeps_internal_hyperlinks_pointing_at_it(ExcelFileFormat format) {
        using var seed = ExcelDocument.Create(); var links = seed.AddWorksheet("Links"); var target = seed.AddWorksheet("Target");
        target.CellValue(2, 2, "Destination"); links.SetInternalLink(1, 1, target, "B2", "Jump", style: false);
        using var document = ExcelDocument.Load(new MemoryStream(seed.ToBytes(format)));
        document.GetSheet("Target").Name = "Sales Łódź";
        byte[] bytes = document.ToBytes(format);
        using var loaded = ExcelDocument.Load(new MemoryStream(bytes));
        var link = Assert.Single(loaded.GetSheet("Links").WorksheetPart.Worksheet!.Descendants<Hyperlink>());
        Assert.Equal("'Sales Łódź'!B2", link.Location?.Value);
        document.GetSheet("Links").CellValue(2, 1, "Second save");
        using var second = ExcelDocument.Load(new MemoryStream(document.ToBytes(format)));
        Assert.Equal(link.Location?.Value, Assert.Single(second.GetSheet("Links").WorksheetPart.Worksheet!.Descendants<Hyperlink>()).Location?.Value);
    }

    private static OfficeVbaProject ReadProject(OfficeCompoundFile source) {
        var storage = OfficeCompoundStorage.Extract(source, "_VBA_PROJECT_CUR");
        return OfficeVbaProject.Load(OfficeCompoundFileWriter.Rewrite(storage, new Dictionary<string, byte[]>()));
    }
}
