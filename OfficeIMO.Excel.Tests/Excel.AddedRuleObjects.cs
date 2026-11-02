using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelAddedRuleObjectTests {
    [Fact]
    public void ReturnedConditionalRulesRetainTheirOwnIdentityThroughLaterAddsAndUpdates() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Values");
        sheet.CellValue(1, 1, "Amount");
        sheet.CellValue(2, 1, 10);
        var first = sheet.AddConditionalRule("A2:A4", ExcelConditionalFormattingOperator.GreaterThan, "5");
        var second = sheet.AddConditionalFormulaRule("A2:A4", "A2<2");
        var scale = sheet.AddConditionalColorScale("B2:B4", "FF0000", "00FF00");
        var bar = sheet.AddConditionalDataBar("C2:C4", "0000FF");
        var icons = sheet.AddConditionalIconSet("D2:D4");
        Assert.Equal("ColorScale", scale.Type);
        Assert.Equal("DataBar", bar.Type);
        Assert.Equal("IconSet", icons.Type);
        first.Range = "A2:A8";
        first.Formulas = new[] { "7" };
        first.StopIfTrue = true;
        sheet.UpdateConditionalFormattingRule(first);
        sheet.RemoveConditionalFormattingRule(second);
        Assert.Equal(4, sheet.GetConditionalFormattingRules().Count);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using (var package = SpreadsheetDocument.Open(output, false))
            Assert.Empty(new OpenXmlValidator().Validate(package));
        output.Position = 0;
        using var loaded = ExcelDocument.Load(output);
        var restored = loaded.Sheets.Single().GetConditionalFormattingRules().Single(rule => rule.Type == "CellIs");
        Assert.Equal("A2:A8", restored.Range);
        Assert.Equal(new[] { "7" }, restored.Formulas);
        Assert.True(restored.StopIfTrue);
    }

    [Fact]
    public void AutoFilterHandleTracksTableOwnershipAndTocReturnsItsCreatedSheet() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Values");
        sheet.CellValue(1, 1, "Name");
        sheet.CellValue(2, 1, "Alpha");
        var filter = sheet.AddAutoFilter("A1:A2");
        Assert.Same(sheet, filter.Sheet);
        Assert.False(filter.Inspect()!.IsTableFilter);
        sheet.AddTable("A1:A2", name: "ValuesTable");
        Assert.Equal("ValuesTable", filter.Inspect()!.TableName);
        var contents = document.AddTableOfContents(styled: false);
        Assert.Equal(contents.Name, document.Sheets[0].Name);
        contents.CellValue(1, 1, "Edited navigation");
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var loaded = ExcelDocument.Load(output);
        Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out var title));
        Assert.Equal("Edited navigation", title);
    }
}
