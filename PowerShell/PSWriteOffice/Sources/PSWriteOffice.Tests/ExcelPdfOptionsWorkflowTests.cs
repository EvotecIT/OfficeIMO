using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text.Json;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelPdfOptionsWorkflowTests {
    [Fact]
    public void PdfFactoryCapturesLiveSheetsByNameForReusableNativeConfiguration() {
        using var original = ExcelDocument.Create();
        var live = original.AddSheet("Hidden Inputs");
        original.AddSheet("Summary");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("live", live);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("New-OfficeExcelPdfOptions -Sheet @($live, 2) -PageSize A4 -UseWorksheetCellStyles:$false -RespectWorkbookSheetVisibility");
        var options = Assert.IsType<ExcelPdfExportOptions>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(PdfPageSizes.A4.Width, Assert.IsType<PdfPageSize>(options.PageSize).Width);
        Assert.False(options.UseWorksheetCellStyles);
        Assert.True(options.RespectWorkbookSheetVisibility);
        Assert.NotNull(options.Sheets);
        Assert.Equal(2, options.Sheets.Count);
        Assert.Equal("Hidden Inputs", options.Sheets[0].Name);
        Assert.Null(options.Sheets[0].Sheet);
        Assert.Equal(2, options.Sheets[1].Position);
        using var reloaded = ExcelDocument.Create();
        reloaded.AddSheet("Hidden Inputs").SetHidden(true);
        reloaded.AddSheet("Summary");
        var persisted = JsonSerializer.Deserialize<OfficeSheetSelector[]>(JsonSerializer.Serialize(options.Sheets));
        Assert.Equal(new[] { "Hidden Inputs", "Summary" }, OfficeSheetSelector.ResolveMany(reloaded, persisted).Select(sheet => sheet.Name));

        shell.Commands.Clear();
        shell.AddCommand("New-OfficeExcelPdfOptions");
        var defaults = Assert.IsType<ExcelPdfExportOptions>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.Equal(new ExcelPdfExportOptions().UseWorksheetCellStyles, defaults.UseWorksheetCellStyles);
        Assert.Null(defaults.Sheets);
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeExcelPdfOptions", typeof(NewOfficeExcelPdfOptionsCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
