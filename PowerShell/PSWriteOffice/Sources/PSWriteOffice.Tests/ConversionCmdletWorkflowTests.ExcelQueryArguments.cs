using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData("Get-OfficeExcelFormula", "Sheet", "Missing")]
    [InlineData("Get-OfficeExcelFormula", "Sheet", "[z-a]")]
    [InlineData("Get-OfficeExcelValidation", "Range", "A1::A2")]
    [InlineData("Get-OfficeExcelValidation", "Header", "Missing")]
    [InlineData("Get-OfficeExcelConditionalFormat", "Range", "A1::A2")]
    public void ExcelQueryInvalidArgumentsTerminateForFilesAndBorrowedDocuments(string command, string parameter, string value) {
        string root = Directory.CreateTempSubdirectory("officeimo-query-arguments-").FullName;
        try {
            string path = Path.Combine(root, "source.xlsx");
            using var document = ExcelDocument.Create();
            document.AddSheet("Data").CellValue(1, 1, "Original");
            document.Save(path);
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("document", document);
            runspace.SessionStateProxy.SetVariable("path", path);
            runspace.SessionStateProxy.SetVariable("query", command);
            runspace.SessionStateProxy.SetVariable("parameter", parameter);
            runspace.SessionStateProxy.SetVariable("value", value);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            foreach (string source in new[] { "LiteralPath", "Document" }) {
                runspace.SessionStateProxy.SetVariable("source", source);
                shell.Commands.Clear(); shell.Streams.Error.Clear();
                shell.AddScript("""
                    $arguments = @{ $parameter = $value }
                    if ($source -eq 'LiteralPath') { $arguments.LiteralPath = $path }
                    else { $arguments.Document = $document }
                    try {
                        & $query @arguments
                        [pscustomobject]@{ Caught = $false; Category = $null }
                    } catch {
                        [pscustomobject]@{ Caught = $true; Category = $_.CategoryInfo.Category }
                    }
                    """);
                var result = Assert.Single(shell.Invoke());
                Assert.True((bool) result.Properties["Caught"].Value, $"{command} -{source} must terminate for invalid -{parameter}.");
                Assert.Equal(ErrorCategory.InvalidArgument, result.Properties["Category"].Value);
            }
        } finally { Directory.Delete(root, true); }
    }
}
