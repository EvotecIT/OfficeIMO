using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Fact]
    public void FormulaQueryReturnsNativeRecordsWithConjunctiveFiltersAndMixedSheetSelectors() {
        using var document = ExcelDocument.Create();
        var summary = document.AddSheet("Summary");
        summary.CellFormula(1, 1, "SUM(1,2)");
        summary.CellFormula(2, 1, "NOW()");
        summary.CellFormula(3, 1, "'[External.xlsx]Data'!A1");
        summary.CellFormula(4, 1, "NOW()+'[External.xlsx]Data'!A1");
        document.AddSheet("Other").CellFormula(1, 1, "TODAY()");
        document.AddSheet("2").CellFormula(1, 1, "COUNT(1,2)");
        using var runspace = OpenRunspace();
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddCommand("Get-OfficeExcelFormula").AddParameter("Document", document);
        var all = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(6, all.Count);
        Assert.All(all, item => Assert.IsType<ExcelFormulaInfo>(item.BaseObject));
        shell.Commands.Clear();
        shell.AddCommand("Get-OfficeExcelFormula").AddParameter("Document", document)
            .AddParameter("Sheet", "Summary").AddParameter("Volatile").AddParameter("External");
        var filtered = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        var both = Assert.IsType<ExcelFormulaInfo>(Assert.Single(filtered).BaseObject);
        Assert.Equal("A4", both.Address);
        Assert.True(both.IsVolatile); Assert.True(both.HasExternalReference);
        shell.Commands.Clear();
        shell.AddCommand("Get-OfficeExcelFormula").AddParameter("Document", document)
            .AddParameter("Sheet", new object[] { 2, "2", summary, "S*" });
        var selected = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(6, selected.Count);
        Assert.Equal(3, selected.Select(item => Assert.IsType<ExcelFormulaInfo>(item.BaseObject).SheetName).Distinct().Count());
        summary.CellFormula(5, 1, "2+2");
        Assert.Equal(7, document.AnalyzeFormulas().FormulaCount);
        string scopedRoot = Directory.CreateTempSubdirectory("officeimo-formula-scope-").FullName;
        try {
            runspace.SessionStateProxy.SetVariable("scopePath", Path.Combine(scopedRoot, "scope.xlsx"));
            shell.Commands.Clear();
            shell.AddScript("""
                New-OfficeExcel -Path $scopePath {
                    ExcelSheet 'Scoped' {
                        ExcelCell -Address A1 -Formula 'SUM(1,2)'
                        $script:scopedFormulas = @(Get-OfficeExcelFormula)
                    }
                }
                $script:scopedFormulas
                """);
            var scoped = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal("Scoped", Assert.IsType<ExcelFormulaInfo>(Assert.Single(scoped).BaseObject).SheetName);
        } finally { Directory.Delete(scopedRoot, true); }
    }

    [Fact]
    public void FormulaQueryReadsLiteralFileInfoFromAnotherDirectoryAndContinuesAfterMissingInputs() {
        string root = Directory.CreateTempSubdirectory("officeimo-formula-query-").FullName;
        string literal = Path.Combine(root, "audit[1].xlsx"), wildcardMatch = Path.Combine(root, "audit1.xlsx");
        try {
            foreach (var item in new[] { (Path: literal, Formula: "NOW()"), (Path: wildcardMatch, Formula: "SUM(1,2)") }) {
                using var document = ExcelDocument.Create(); document.AddSheet("Data").CellFormula(1, 1, item.Formula); document.Save(item.Path);
            }
            byte[] original = File.ReadAllBytes(literal);
            string elsewhere = Path.Combine(root, "elsewhere"); Directory.CreateDirectory(elsewhere);
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("literal", literal);
            runspace.SessionStateProxy.SetVariable("root", root);
            runspace.SessionStateProxy.SetVariable("elsewhere", elsewhere);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("Set-Location -LiteralPath $elsewhere; Get-Item -LiteralPath $literal | Get-OfficeExcelFormula -Volatile");
            var result = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal("NOW()", Assert.IsType<ExcelFormulaInfo>(Assert.Single(result).BaseObject).Formula);
            Assert.Equal(literal, result[0].Properties["Path"].Value);
            shell.Commands.Clear();
            shell.AddScript("Get-OfficeExcelFormula -Path (Join-Path $root 'audit*.xlsx')");
            var expanded = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(2, expanded.Count);
            shell.Commands.Clear();
            string missing = Path.Combine(root, "missing.xlsx");
            shell.AddCommand("Get-OfficeExcelFormula").AddParameter("LiteralPath", new[] { missing, literal })
                .AddParameter("ErrorAction", ActionPreference.Continue);
            var continued = shell.Invoke();
            Assert.True(shell.HadErrors);
            Assert.Single(continued);
            Assert.Equal(missing, Assert.Single(shell.Streams.Error).TargetObject);
            Assert.Equal(original, File.ReadAllBytes(literal));
            using var exclusive = File.Open(literal, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        } finally { Directory.Delete(root, true); }
    }
}
