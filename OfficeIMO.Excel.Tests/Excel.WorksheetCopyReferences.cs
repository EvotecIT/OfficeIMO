using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SameWorkbookCopiesRebindSelfReferencesAcrossCopiedMetadata(bool copyFrom) {
        using var document = ExcelDocument.Create();
        ExcelSheet source = document.AddSheet("Original");
        document.AddSheet("Other").Cell(1, 1).SetValue(11);
        source.Cell(1, 1).SetValue(100);
        source.Cell(2, 1).SetFormula("Original!A1*3");
        const string otherReferences = "Other!A1+'[Other.xlsx]Original'!A1+[Other.xlsx]Original!A1";
        source.Cell(2, 2).SetFormula(otherReferences);
        source.Cell(2, 3).SetFormula("IF(TRUE,\"Original!A1\",'Original'!A1)");
        source.Cell(2, 4).SetFormula("SUM(Sales[Amount])+SUM(Sales2[Amount])");
        source.Cell(2, 5).SetFormula("LocalAnchor*3");
        document.SetNamedRange("LocalAnchor", "A1", source, save: false);
        source.Range("E3").Validation().CustomFormula("Original!A1>0");
        source.Range("F3").ConditionalFormat().CellIs(ExcelConditionalFormattingOperator.GreaterThan, "Original!$A$1");
        source.Cell(4, 1).SetValue("Amount");
        source.Cell(4, 2).SetValue("Value");
        source.Cell(5, 1).SetValue(10);
        source.Cell(5, 2).SetValue(100);
        source.AddTable("A4:B5", true, "Sales");
        source.Cell(7, 1).SetValue("Amount");
        source.Cell(8, 1).SetValue(20);
        source.AddTable("A7:A8", true, "Sales2");
        Table originalTable = source.WorksheetPart.TableDefinitionParts.Single(part => part.Table!.Name == "Sales").Table!;
        TableColumn valueColumn = originalTable.TableColumns!.Elements<TableColumn>().Last();
        valueColumn.CalculatedColumnFormula = new CalculatedColumnFormula("Original!A1+SUM(Sales[Amount])");
        valueColumn.TotalsRowFormula = new TotalsRowFormula("SUM(Original!A1,Sales2[Amount])");
        originalTable.Save();
        source.WorksheetPart.Worksheet.AddChild(new Hyperlinks(new Hyperlink { Reference = "G1", Location = "Original!A1" }), true);
        var external = source.WorksheetPart.AddHyperlinkRelationship(new Uri("Other.xlsx", UriKind.Relative), true);
        source.WorksheetPart.Worksheet.GetFirstChild<Hyperlinks>()!.Append(new Hyperlink {
            Reference = "G2", Id = external.Id, Location = "Original!A1"
        });

        ExcelSheet copied = copyFrom
            ? document.CopyWorksheetFrom(document, "Original", "Copy:O'Brien", ExcelSheetNameValidationMode.Sanitize,
                new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package })
            : document.CopyWorksheet(source, "Copy:O'Brien");
        Assert.Equal("Copy_O'Brien", copied.Name);
        copied.Cell(1, 1).SetValue(7);
        using var output = new MemoryStream();
        document.Save(output, new ExcelSaveOptions { EvaluateFormulasBeforeSave = true });
        output.Position = 0;
        using var package = SpreadsheetDocument.Open(output, false);
        WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(package, "Copy_O'Brien");
        var cells = copiedPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
        var links = copiedPart.Worksheet.Descendants<Hyperlink>().ToDictionary(link => link.Reference!.Value!);
        Assert.Equal("'Copy_O''Brien'!A1", links["G1"].Location!.Value);
        Assert.Equal("Original!A1", links["G2"].Location!.Value);
        Assert.Equal("Other.xlsx", copiedPart.HyperlinkRelationships.Single(link => link.Id == links["G2"].Id!.Value).Uri.OriginalString);
        Assert.Equal("'Copy_O''Brien'!A1*3", cells["A2"].CellFormula!.Text);
        Assert.Equal("21", cells["A2"].CellValue!.Text);
        Assert.Equal(otherReferences, cells["B2"].CellFormula!.Text);
        Assert.Equal("IF(TRUE,\"Original!A1\",'Copy_O''Brien'!A1)", cells["C2"].CellFormula!.Text);
        Assert.Equal("SUM(Sales3[Amount])+SUM(Sales22[Amount])", cells["D2"].CellFormula!.Text);
        Assert.Equal("21", cells["E2"].CellValue!.Text);
        Assert.Equal("'Copy_O''Brien'!A1>0", copiedPart.Worksheet.Descendants<Formula1>().Single().Text);
        Assert.Equal("'Copy_O''Brien'!$A$1", copiedPart.Worksheet.Descendants<ConditionalFormattingRule>().Single().Elements<Formula>().Single().Text);
        Table copiedTable = copiedPart.TableDefinitionParts.Single(part => part.Table!.Name == "Sales3").Table!;
        TableColumn copiedValue = copiedTable.TableColumns!.Elements<TableColumn>().Last();
        Assert.Equal("'Copy_O''Brien'!A1+SUM(Sales3[Amount])", copiedValue.CalculatedColumnFormula!.Text);
        Assert.Equal("SUM('Copy_O''Brien'!A1,Sales22[Amount])", copiedValue.TotalsRowFormula!.Text);
        DefinedName localName = package.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>()
            .Single(name => name.Name == "LocalAnchor" && name.LocalSheetId!.Value == 2U);
        Assert.Equal("'Copy_O''Brien'!$A$1", localName.Text);
        WorksheetPart sourcePart = GetWorksheetPartByNameForOperations(package, "Original");
        Cell originalFormula = sourcePart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "A2");
        Assert.Equal("Original!A1*3", originalFormula.CellFormula!.Text);
        Assert.Equal("300", originalFormula.CellValue!.Text);
        Assert.Equal("Original!A1+SUM(Sales[Amount])", sourcePart.TableDefinitionParts
            .Single(part => part.Table!.Name == "Sales").Table!.Descendants<CalculatedColumnFormula>().Single().Text);
    }
}
