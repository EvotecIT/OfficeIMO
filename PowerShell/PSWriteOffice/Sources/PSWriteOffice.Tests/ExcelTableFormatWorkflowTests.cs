using System.Collections;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelTableFormatWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NamedTableFormatsOnlyItsOwnDataCells(bool noHeader) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Sales");
        sheet.CellValue(20, 2, 0.25);
        sheet.FormatCell(20, 2, "0.0000");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("sheet", sheet);
        runspace.SessionStateProxy.SetVariable("noHeader", noHeader);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            Add-OfficeExcelTable -Worksheet $sheet -InputObject @(
                [pscustomobject]@{ Region = 'North'; Growth = 0.12 }
                [pscustomobject]@{ Region = 'South'; Growth = -0.03 }
            ) -Name Sales -ColumnFormat @{ Growth = 'Percent' } -NoHeader:$noHeader
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var saved = ExcelDocument.Load(output);
        var reopened = saved.Sheets[0];
        Assert.NotNull(reopened.GetTableRange("Sales"));
        int firstDataRow = noHeader ? 1 : 2;
        Assert.Equal("0.00%", reopened.CellAt(firstDataRow, 2).GetStyle().NumberFormatCode);
        Assert.Equal("0.0000", reopened.CellAt(20, 2).GetStyle().NumberFormatCode);
        Assert.True(reopened.TryGetCellText(firstDataRow, 1, out string region));
        Assert.Equal("North", region);
    }

    [Fact]
    public void MissingFormatColumnRejectsTheTableBeforeWriting() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Sales");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Add-OfficeExcelTable").AddParameter("Worksheet", sheet)
            .AddParameter("InputObject", new Hashtable { ["Growth"] = 0.12 })
            .AddParameter("ColumnFormat", new Hashtable { ["Typo"] = "Percent" });
        var failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.False(sheet.TryGetCellText(1, 1, out _));
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeExcelTable", typeof(AddOfficeExcelTableCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
