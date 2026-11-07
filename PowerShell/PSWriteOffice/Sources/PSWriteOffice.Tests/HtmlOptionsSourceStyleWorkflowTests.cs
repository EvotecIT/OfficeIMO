using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel.Html;
using OfficeIMO.PowerPoint.Html;
using PSWriteOffice.Cmdlets.Excel;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class HtmlOptionsSourceStyleWorkflowTests {
    [Theory]
    [InlineData("New-OfficeExcelHtmlOptions")]
    [InlineData("New-OfficePowerPointHtmlOptions")]
    public void HtmlOptionsKeepSourceStylesUnlessPresentationIsRequested(string command) {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeExcelHtmlOptions", typeof(NewOfficeExcelHtmlOptionsCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficePowerPointHtmlOptions", typeof(NewOfficePowerPointHtmlOptionsCommand), null));
        using var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        foreach (var (arguments, expectedStyles) in new[] { ("", false), (" -Theme Modern", true), (" -Theme Modern -NoDefaultStyles", false) }) {
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript(command + arguments);
            object options = Assert.Single(shell.Invoke()).BaseObject;
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            bool includeStyles = options switch {
                ExcelHtmlExportOptions excel => excel.DocumentOutput.IncludeDefaultStyles,
                PowerPointHtmlExportOptions deck => deck.DocumentOutput.IncludeDefaultStyles,
                _ => throw new InvalidOperationException("Unexpected HTML options type.")
            };
            Assert.Equal(expectedStyles, includeStyles);
        }
    }
}
