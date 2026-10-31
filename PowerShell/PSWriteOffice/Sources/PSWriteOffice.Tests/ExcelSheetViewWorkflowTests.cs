using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;
using PSWriteOffice.Models.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelSheetViewWorkflowTests {
    [Fact]
    public void MergedViewSettingsPreserveOmissionsAndPersistExplicitFalseAndUnfreeze() {
        using var book = ExcelDocument.Create();
        book.AddSheet("Summary");
        var data = book.AddSheet("Data");
        using var runspace = Open(book);
        Assert.Empty(Invoke(runspace, "$book | Set-OfficeExcelSheetView -Sheet 2 -FreezeRows 3 -FreezeColumns 1 -NoGridlines -RightToLeft -Zoom 85 -ZoomNormal 100 -View PageLayout"));
        var output = Invoke(runspace, "Set-OfficeExcelSheetView -Sheet $selected -FreezeRows 4 -PassThru");
        Assert.Same(book["Data"], Assert.Single(output).BaseObject);
        var view = data.GetViewInfo();
        Assert.Equal(4, view.FrozenRowCount);
        Assert.Equal(1, view.FrozenColumnCount);
        Assert.False(view.ShowGridlines);
        Assert.True(view.RightToLeft);
        Assert.Equal(85U, view.ZoomScale);
        Assert.Equal("pageLayout", view.View);
        Assert.False(book["Summary"].GetViewInfo().HasPane);
        Assert.Empty(Invoke(runspace, "Set-OfficeExcelSheetView -Document $book -Sheet Data -FreezeRows 0 -FreezeColumns 0 -NoGridlines:$false -RightToLeft:$false"));
        using var bytes = new MemoryStream();
        book.Save(bytes);
        bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        var persisted = reopened["Data"].GetViewInfo();
        Assert.False(persisted.HasPane);
        Assert.True(persisted.ShowGridlines);
        Assert.False(persisted.RightToLeft);
        Assert.Equal(85U, persisted.ZoomScale);
        Assert.Equal(100U, persisted.ZoomScaleNormal);
        Assert.Equal("pageLayout", persisted.View);
    }

    [Theory]
    [InlineData("-Sheet Data -View 99 -FreezeRows 2 -NoGridlines")]
    [InlineData("-Sheet Data -Zoom 1 -FreezeRows 2 -NoGridlines")]
    [InlineData("-Sheet Data -FreezeRows -1 -NoGridlines")]
    [InlineData("-Sheet $foreign -FreezeRows 2 -NoGridlines")]
    public void InvalidViewArgumentsFailBeforeAnyWorksheetChange(string arguments) {
        using var book = ExcelDocument.Create();
        book.AddSheet("Summary");
        var data = book.AddSheet("Data");
        data.Freeze(1, 1);
        using var foreign = ExcelDocument.Create();
        using var runspace = Open(book);
        runspace.SessionStateProxy.SetVariable("foreign", foreign.AddSheet("Data"));
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Set-OfficeExcelSheetView -Document $book " + arguments);
        var error = Record.Exception(() => shell.Invoke());
        Assert.True(error != null || shell.HadErrors);
        var view = data.GetViewInfo();
        Assert.True(view.ShowGridlines);
        Assert.Equal(1, view.FrozenRowCount);
        Assert.Equal(1, view.FrozenColumnCount);
        Assert.Null(view.ZoomScale);
        Assert.Null(view.View);
        Assert.False(book["Summary"].GetViewInfo().HasPane);
    }

    [Fact]
    public void TypedViewInspectionBindsSelectorsAndLiteralFileInfoWithoutWriting() {
        string root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PSWriteOffice-view-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using var book = ExcelDocument.Create();
            book.AddSheet("Summary");
            var data = book.AddSheet("Data");
            data.Freeze(2, 1);
            using var runspace = Open(book);
            var mixed = Invoke(runspace, "$book | Get-OfficeExcelSheetView -Sheet 'Summ*', 2, $selected");
            Assert.Equal(new[] { "Summary", "Data" }, mixed.Select(item => Assert.IsType<ExcelSheetViewInfo>(item.BaseObject).SheetName));
            var inferred = Assert.IsType<ExcelSheetViewInfo>(Assert.Single(Invoke(runspace, "Get-OfficeExcelSheetView -Sheet $selected")).BaseObject);
            Assert.Equal(2, inferred.FrozenRowCount);
            Assert.Null(inferred.Path);
            string path = System.IO.Path.Combine(root, "report[1].xlsx");
            book.Save(path);
            byte[] before = File.ReadAllBytes(path);
            runspace.SessionStateProxy.SetVariable("file", new FileInfo(path));
            runspace.SessionStateProxy.SetVariable("literal", path);
            runspace.SessionStateProxy.SetVariable("pattern", System.IO.Path.Combine(root, "report*.xlsx"));
            foreach (string input in new[] { "$file | Get-OfficeExcelSheetView -Sheet Data", "Get-OfficeExcelSheetView -LiteralPath $literal -Sheet 2", "Get-OfficeExcelSheetView -Path $pattern -Sheet 'D*'" }) {
                var view = Assert.IsType<ExcelSheetViewInfo>(Assert.Single(Invoke(runspace, input)).BaseObject);
                Assert.Equal("Data", view.SheetName);
                Assert.Equal(path, view.Path);
                Assert.Equal(2, view.FrozenRowCount);
                Assert.Equal(1, view.FrozenColumnCount);
            }
            Assert.Equal(before, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static Runspace Open(ExcelDocument book) {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Set-OfficeExcelSheetView", typeof(SetOfficeExcelSheetViewCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Get-OfficeExcelSheetView", typeof(GetOfficeExcelSheetViewCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        runspace.SessionStateProxy.SetVariable("book", book);
        runspace.SessionStateProxy.SetVariable("selected", book["Data"]);
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
