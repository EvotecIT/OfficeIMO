using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelNativeEditingReadbacks {
    private static string Corpus => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus");

    [Theory]
    [InlineData("xls")]
    [InlineData("xlsb")]
    public void Native_producer_retains_worksheet_order_calculation_and_document_hosts(string extension) {
        string path = Path.Combine(Corpus, "excel-lifecycle-readback." + extension);
        using var document = ExcelDocument.Load(path);
        Assert.Equal(new[] { "b_chart", "b", "lb_chart", "lb", "c_chart", "c", "Info", "cl_chart", "cl", "maquis", "wijn", "All", "Sales Łódź", "Added data" },
            document.WorkbookRoot.Sheets!.ChildElements.Select(sheet => sheet.GetAttribute("name", "").Value));
        Assert.Equal(42D, document.GetSheet("Sales Łódź").CellAt(1, 1).GetValue<double>());
        Assert.Equal(43D, document.GetSheet("Added data").CellAt(2, 1).GetValue<double>());
        Assert.Equal("DataCode1", document.GetSheet("Sales Łódź").WorksheetPart.Worksheet!.SheetProperties?.CodeName?.Value);
        string? addedCodeName = document.GetSheet("Added data").WorksheetPart.Worksheet!.SheetProperties?.CodeName?.Value;
        Assert.False(string.IsNullOrEmpty(addedCodeName));
        OfficeVbaProject project;
        if (extension == "xls") {
            var native = LegacyXlsWorkbook.Load(path); Assert.Equal(4, native.ChartSheets.Count);
            var storage = OfficeCompoundStorage.Extract(native.SourceCompoundFile!, "_VBA_PROJECT_CUR");
            project = OfficeVbaProject.Load(OfficeCompoundFileWriter.Rewrite(storage, new Dictionary<string, byte[]>()));
        } else { Assert.Equal(4, document.WorkbookPartRoot.ChartsheetParts.Count()); project = document.ReadVbaProject()!; }
        Assert.DoesNotContain(project.Modules, module => module.Name == "DataCode2");
        Assert.Contains(project.Modules, module => module.Name == "DataCode1");
        Assert.Contains(project.Modules, module => module.Name == addedCodeName && module.Kind == OfficeVbaModuleKind.Document);
    }

    [Theory]
    [InlineData("xls")]
    [InlineData("xlsb")]
    public void Native_producer_recognizes_multicell_array_after_cache_clear(string extension) {
        using var document = ExcelDocument.Load(Path.Combine(Corpus, "excel-array-readback." + extension));
        var sheet = document.Sheets[0];
        Assert.Equal(5D, sheet.CellAt(1, 2).GetValue<double>()); Assert.Equal(5D, sheet.CellAt(2, 2).GetValue<double>());
        var owner = sheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single(cell => cell.CellReference?.Value == "B1");
        Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Array, owner.CellFormula?.FormulaType?.Value);
        Assert.Equal("B1:B2", owner.CellFormula?.Reference?.Value); Assert.Equal("SUM(A1:A2)", owner.CellFormula?.Text);
        Assert.True(sheet.CellAt(2, 2).GetStyle().Bold);
    }
}
