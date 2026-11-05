using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO;
using OfficeIMO.Excel.Html;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelHtmlOptionsWorkflowTests {
    [Fact]
    public void HtmlFactoryForwardsTheSharedThemeAndBoundStyleSwitchToNativeOutput() {
        using var runspace = Open();
        var theme = OfficeTheme.Modern;
        runspace.SessionStateProxy.SetVariable("theme", theme);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("New-OfficeExcelHtmlOptions -Profile VisualReview -Theme $theme -Title 'Operations' -NoDefaultStyles:$false");
        var options = Assert.IsType<ExcelHtmlExportOptions>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(ExcelHtmlExportProfile.VisualReview, options.ExportProfile);
        Assert.Same(theme, options.DocumentOutput.DocumentTheme);
        Assert.Equal("Operations", options.DocumentOutput.Title);
        Assert.True(options.DocumentOutput.IncludeDefaultStyles);
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeExcelHtmlOptions", typeof(NewOfficeExcelHtmlOptionsCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
