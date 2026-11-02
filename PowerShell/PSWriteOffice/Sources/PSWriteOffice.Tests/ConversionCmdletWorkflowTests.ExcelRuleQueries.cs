using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData("Get-OfficeExcelValidation")]
    [InlineData("Get-OfficeExcelConditionalFormat")]
    public void RuleQueriesPreferHeadersToRangesAndRetainNativeSnapshotsAndWorksheetProvenance(string command) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        sheet.CellValue(1, 1, "Status"); sheet.CellValue(1, 2, "Value");
        sheet.CellValue(2, 1, "Open"); sheet.CellValue(3, 1, "Closed");
        sheet.CellValue(2, 2, 10); sheet.CellValue(3, 2, 20);
        sheet.ValidationList("A2:A3", new[] { "Open", "Closed" });
        sheet.ValidationList("B2:B3", new[] { "10", "20" });
        sheet.AddConditionalRule("A2:A3", ExcelConditionalFormattingOperator.Equal, "\"Open\"");
        sheet.AddConditionalRule("B2:B3", ExcelConditionalFormattingOperator.GreaterThan, "10");
        using var runspace = OpenRunspace();
        runspace.SessionStateProxy.SetVariable("document", document);
        runspace.SessionStateProxy.SetVariable("sheet", sheet);
        runspace.SessionStateProxy.SetVariable("query", command);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            $rule = & $query -Document $document -Sheet Data -Range B2 -Header Status
            if ($query -eq 'Get-OfficeExcelConditionalFormat') {
                $rule.StopIfTrue = $true
                $sheet.UpdateConditionalFormattingRule($rule)
            } else { $rule.PromptTitle = 'Snapshot only' }
            $rule
            """);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        var item = Assert.Single(output);
        Assert.Equal("Data", item.Properties["SheetName"].Value);
        Assert.Equal("Data", item.Properties["Sheet"].Value);
        if (command == "Get-OfficeExcelValidation") {
            var rule = Assert.IsType<ExcelDataValidationInfo>(item.BaseObject);
            Assert.Equal("A2:A3", rule.Range);
            Assert.Contains("Open", rule.Formula1, StringComparison.Ordinal);
            Assert.Null(Assert.Single(sheet.GetDataValidations("A2")).PromptTitle);
        } else {
            var rule = Assert.IsType<ExcelConditionalFormattingInfo>(item.BaseObject);
            Assert.Equal("A2:A3", rule.Range);
            Assert.True(Assert.Single(sheet.GetConditionalFormattingRules("A2")).StopIfTrue);
        }
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(2, reopened.Sheets[0].GetDataValidations().Count);
        Assert.Equal(2, reopened.Sheets[0].GetConditionalFormattingRules().Count);
        Assert.Equal(command == "Get-OfficeExcelConditionalFormat",
            Assert.Single(reopened.Sheets[0].GetConditionalFormattingRules("A2")).StopIfTrue);
    }
}
