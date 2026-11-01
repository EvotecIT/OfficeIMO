using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelSheetCopyMoveWorkflowTests {
    [Fact]
    public void NativeSourcesKeepSelectedSheetReferencesAndMoveUsesOneBasedPositions() {
        using var target = ExcelDocument.Create();
        target.AddSheet("Summary").CellValue(1, 1, "Keep");
        using var source = ExcelDocument.Create();
        source.AddSheet("Sales West").CellValue(1, 1, 42);
        source.AddSheet("Sales Total").CellFormula(1, 1, "'Sales West'!A1");
        source.AddSheet("Other").CellValue(1, 1, "Do not import");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("book", target);
        runspace.SessionStateProxy.SetVariable("source", PSObject.AsPSObject(source));

        var copied = Invoke(runspace, "$book | Copy-OfficeExcelSheet -Source $source -Sheet 'Sales T*',1 -NamePrefix 'Imported ' -PassThru -Confirm:$false")
            .Select(item => Assert.IsType<ExcelSheet>(item.BaseObject)).ToArray();
        Assert.Equal(new[] { "Imported Sales West", "Imported Sales Total" }, copied.Select(sheet => sheet.Name));
        Assert.All(copied, sheet => Assert.Same(target, sheet.Document));
        Assert.Contains("Imported Sales West", copied[1].CellAt(1, 1).GetValue().Formula);
        Assert.Equal(42, copied[0].CellAt(1, 1).GetValue<int>());
        Assert.Equal("Keep", target["Summary"].CellAt(1, 1).GetValue<string>());
        Assert.Equal(3, source.Sheets.Count);
        runspace.SessionStateProxy.SetVariable("copied", copied[1]);
        Assert.Empty(Invoke(runspace, "$book | Move-OfficeExcelSheet -Sheet $copied -Position 1 -WhatIf"));
        Assert.Equal("Summary", target.Sheets[0].Name);
        Assert.Same(copied[1], Assert.Single(Invoke(runspace,
            "$book | Move-OfficeExcelSheet -Sheet $copied -Position 1 -PassThru -Confirm:$false")).BaseObject);

        using var saved = new MemoryStream();
        target.Save(saved);
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal("Imported Sales Total", reopened.Sheets[0].Name);
        Assert.Contains("Imported Sales West", reopened.Sheets[0].CellAt(1, 1).GetValue().Formula);
        Assert.Equal(42, reopened["Imported Sales West"].CellAt(1, 1).GetValue<int>());
        Assert.Equal(42, source["Sales West"].CellAt(1, 1).GetValue<int>());
    }

    [Theory]
    [InlineData("$file | Copy-OfficeExcelSheet")]
    [InlineData("Copy-OfficeExcelSheet -LiteralPath $literal")]
    [InlineData("Copy-OfficeExcelSheet -Path $pattern")]
    public void FileCopiesPersistAndReturnedNativeSheetsKeepTheirWorkbookOpen(string command) {
        string root = Path.Combine(Path.GetTempPath(), "PSWriteOffice-copy-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string targetPath = Path.Combine(root, "target[1].xlsx");
            string sourcePath = Path.Combine(root, "source[1].xlsx");
            using (var target = ExcelDocument.Create()) {
                target.AddSheet("Summary").CellValue(1, 1, "Keep");
                target.Save(targetPath);
            }
            using (var source = ExcelDocument.Create()) {
                source.AddSheet("Incoming").CellValue(1, 1, "Imported");
                source.Save(sourcePath);
            }
            byte[] before = File.ReadAllBytes(targetPath);
            byte[] sourceBefore = File.ReadAllBytes(sourcePath);
            using var runspace = Open();
            runspace.SessionStateProxy.SetVariable("file", new FileInfo(targetPath));
            runspace.SessionStateProxy.SetVariable("literal", targetPath);
            runspace.SessionStateProxy.SetVariable("pattern", Path.Combine(root, "target*.xlsx"));
            runspace.SessionStateProxy.SetVariable("source", new FileInfo(sourcePath));
            string arguments = " -Source $source -Sheet 1 -Name Copied -PassThru";
            Assert.Empty(Invoke(runspace, command + arguments + " -WhatIf"));
            Assert.Equal(before, File.ReadAllBytes(targetPath));

            ExcelSheet copy = Assert.IsType<ExcelSheet>(Assert.Single(Invoke(runspace,
                command + arguments + " -Confirm:$false")).BaseObject);
            using (copy.Document) {
                Assert.Equal("Imported", copy.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Keep", copy.Document["Summary"].CellAt(1, 1).GetValue<string>());
                copy.CellValue(2, 1, "Still open");
                copy.Document.Save();
            }
            using var reopened = ExcelDocument.Load(targetPath);
            Assert.Equal("Imported", reopened["Copied"].CellAt(1, 1).GetValue<string>());
            Assert.Equal("Still open", reopened["Copied"].CellAt(2, 1).GetValue<string>());
            Assert.Equal(sourceBefore, File.ReadAllBytes(sourcePath));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData("Copy-OfficeExcelSheet -Document $book -Source $source -Sheet 'Missing*'")]
    [InlineData("Copy-OfficeExcelSheet -Document $book -Source $source -Sheet '*' -Name Single")]
    [InlineData("Copy-OfficeExcelSheet -Document $book -Source $source -Sheet $foreign")]
    [InlineData("Move-OfficeExcelSheet -Document $book -Sheet 1 -Position 2")]
    public void InvalidSelectionsFailBeforeChangingTheWorkbook(string command) {
        using var target = ExcelDocument.Create();
        var keep = target.AddSheet("Keep");
        keep.CellValue(1, 1, "Unchanged");
        using var source = ExcelDocument.Create();
        source.AddSheet("First");
        source.AddSheet("Second");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("book", target);
        runspace.SessionStateProxy.SetVariable("source", source);
        runspace.SessionStateProxy.SetVariable("foreign", keep);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(command + " -Confirm:$false");
        Exception? error = Record.Exception(() => shell.Invoke());
        Assert.True(error != null || shell.HadErrors);
        Assert.Equal("Keep", Assert.Single(target.Sheets).Name);
        Assert.Equal("Unchanged", keep.CellAt(1, 1).GetValue<string>());
        Assert.Equal(2, source.Sheets.Count);
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Copy-OfficeExcelSheet", typeof(CopyOfficeExcelSheetCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Move-OfficeExcelSheet", typeof(MoveOfficeExcelSheetCommand), null));
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
