using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelSheetMergeWorkflowTests {
    [Fact]
    public void LiveSourcesUseTypedSelectorsAndTargetAddressWithoutTakingOwnership() {
        using var target = ExcelDocument.Create();
        target.AddSheet("Other").Cell(1, 1).SetValue("Keep other sheet");
        ExcelSheet combined = target.AddSheet("Combined");
        combined.Cell(4, 2).SetValue("Amount");
        combined.Cell(4, 3).SetValue("Region");
        using var source = ExcelDocument.Create();
        source.AddSheet("Other").Cell(1, 1).SetValue("Do not import");
        ExcelSheet incoming = source.AddSheet("Incoming");
        incoming.Cell(1, 1).SetValue("Region");
        incoming.Cell(1, 2).SetValue("Amount");
        incoming.Cell(2, 1).SetValue("APAC");
        incoming.Cell(2, 2).SetValue(150);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("book", target);
        runspace.SessionStateProxy.SetVariable("source", PSObject.AsPSObject(source));
        runspace.SessionStateProxy.SetVariable("incoming", incoming);

        var output = Invoke(runspace, "$book | Merge-OfficeExcelSheet -Sheet 2 -Source $source -SourceSheet $incoming -Address B5 -HeaderRow 4 -MatchColumnsByHeader -Confirm:$false");
        ExcelWorksheetMergeResult result = Assert.IsType<ExcelWorksheetMergeResult>(Assert.Single(output).BaseObject);
        Assert.Equal("Combined", result.TargetSheetName);
        Assert.Equal("Incoming", result.SourceSheetName);
        Assert.Equal(1, result.RowsCopied);
        Assert.Equal(150, combined.Cell(5, 2).GetValue<int>());
        Assert.Equal("APAC", combined.Cell(5, 3).GetValue<string>());
        Assert.Equal("Keep other sheet", target["Other"].Cell(1, 1).GetValue<string>());
        Assert.Equal("APAC", source["Incoming"].Cell(2, 1).GetValue<string>());
        using var saved = new MemoryStream();
        target.Save(saved);
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(150, reopened["Combined"].Cell(5, 2).GetValue<int>());
        Assert.Equal("APAC", reopened["Combined"].Cell(5, 3).GetValue<string>());
    }

    [Theory]
    [InlineData("$file | Merge-OfficeExcelSheet")]
    [InlineData("Merge-OfficeExcelSheet -LiteralPath $literal")]
    [InlineData("Merge-OfficeExcelSheet -Path $pattern")]
    public void FileTargetsPersistRowsAndWhatIfKeepsBothLiteralWorkbooksUnchanged(string command) {
        string root = Path.Combine(Path.GetTempPath(), "PSWriteOffice-merge-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string targetPath = Path.Combine(root, "target[1].xlsx");
            string sourcePath = Path.Combine(root, "source[1].xlsx");
            using (var target = ExcelDocument.Create()) {
                var data = target.AddSheet("Data");
                data.Cell(1, 1).SetValue("Region");
                data.Cell(2, 1).SetValue("Keep");
                target.Save(targetPath);
            }
            using (var source = ExcelDocument.Create()) {
                var data = source.AddSheet("Incoming");
                data.Cell(1, 1).SetValue("Region");
                data.Cell(2, 1).SetValue("APAC");
                source.Save(sourcePath);
            }
            byte[] targetBefore = File.ReadAllBytes(targetPath);
            byte[] sourceBefore = File.ReadAllBytes(sourcePath);
            using var runspace = Open();
            runspace.SessionStateProxy.SetVariable("file", new FileInfo(targetPath));
            runspace.SessionStateProxy.SetVariable("literal", targetPath);
            runspace.SessionStateProxy.SetVariable("pattern", Path.Combine(root, "target*.xlsx"));
            runspace.SessionStateProxy.SetVariable("source", new FileInfo(sourcePath));
            string arguments = " -Sheet Data -Source $source -SourceSheet 1 -NoSourceHeader:$false";
            Assert.Empty(Invoke(runspace, command + arguments + " -WhatIf"));
            Assert.Equal(targetBefore, File.ReadAllBytes(targetPath));
            Assert.Equal(sourceBefore, File.ReadAllBytes(sourcePath));

            ExcelWorksheetMergeResult result = Assert.IsType<ExcelWorksheetMergeResult>(
                Assert.Single(Invoke(runspace, command + arguments + " -Confirm:$false")).BaseObject);
            Assert.Equal(1, result.RowsCopied);
            using var merged = ExcelDocument.Load(targetPath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            Assert.Equal("Keep", merged["Data"].Cell(2, 1).GetValue<string>());
            Assert.Equal("APAC", merged["Data"].Cell(3, 1).GetValue<string>());
            Assert.Equal(sourceBefore, File.ReadAllBytes(sourcePath));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public void SourceSelectorMustBelongToTheChosenSourceWorkbook() {
        using var target = ExcelDocument.Create();
        var targetSheet = target.AddSheet("Data");
        targetSheet.Cell(1, 1).SetValue("Region");
        targetSheet.Cell(2, 1).SetValue("Keep");
        using var source = ExcelDocument.Create();
        source.AddSheet("Data").Cell(1, 1).SetValue("Wrong source");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("book", target);
        runspace.SessionStateProxy.SetVariable("source", source);
        runspace.SessionStateProxy.SetVariable("foreign", targetSheet);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Merge-OfficeExcelSheet -Document $book -Sheet Data -Source $source -SourceSheet $foreign -Confirm:$false");
        Exception? error = Record.Exception(() => shell.Invoke());
        Assert.True(error != null || shell.HadErrors);
        Assert.Equal("Keep", targetSheet.Cell(2, 1).GetValue<string>());
        Assert.Null(targetSheet.Cell(3, 1).GetValue<string>());
        Assert.Equal("Wrong source", source["Data"].Cell(1, 1).GetValue<string>());
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Merge-OfficeExcelSheet", typeof(MergeOfficeExcelSheetCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static System.Collections.ObjectModel.Collection<PSObject> Invoke(Runspace runspace, string script) {
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        return output;
    }
}
