using System.Collections;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelTextWorkflowTests {
    [Fact]
    public void CellFillAndGradientParametersCompleteSharedThemeRoles() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Set-OfficeExcelCell", typeof(SetOfficeExcelCellCommand), null));
        using var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        foreach (string parameter in new[] { "BackgroundColor", "GradientFrom", "GradientTo" }) {
            string input = $"Set-OfficeExcelCell A1 -{parameter} Da";
            var completion = CommandCompletion.CompleteInput(input, input.Length, new Hashtable(), shell);
            Assert.Contains(completion.CompletionMatches, match => match.ListItemText == "Danger");
        }
    }

    [Fact]
    public void CellChildrenPersistNativeRunFormattingAndExplicitDisabledValuesOnTheSelectedSheet() {
        using var document = ExcelDocument.Create();
        document.AddSheet("2").Cell(1, 1).SetValue("Keep");
        document.AddSheet("Report");
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Set-OfficeExcelCell", typeof(SetOfficeExcelCellCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeExcelText", typeof(AddOfficeExcelTextCommand), null));
        using var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        runspace.SessionStateProxy.SetVariable("book", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            Set-OfficeExcelCell -Document $book -Sheet 2 -Address A1 -Value {
                Add-OfficeExcelText 'Status: ' -Bold -Italic -Strikethrough -Underline Double -Font Arial -Size 14pt -Color Red
                Add-OfficeExcelText 'Ready' -Bold:$false -Italic:$false -Strikethrough:$false -Underline None -VerticalAlignment Superscript
            }
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        using var bytes = new MemoryStream();
        document.Save(bytes);
        bytes.Position = 0;
        using var saved = ExcelDocument.Load(bytes);
        Assert.Equal("Keep", saved["2"].Cell(1, 1).GetValue<string>());
        var runs = saved["Report"].Cell(1, 1).GetRichText();
        Assert.Equal(2, runs.Count);
        Assert.Equal("Status: ", runs[0].Text);
        Assert.True(runs[0].Bold);
        Assert.True(runs[0].Italic);
        Assert.True(runs[0].Strikethrough);
        Assert.Equal(ExcelUnderlineStyle.Double, runs[0].UnderlineStyle);
        Assert.Equal(OfficeColor.Red, runs[0].Color);
        Assert.Equal("Arial", runs[0].FontFamily);
        Assert.Equal(14, runs[0].FontSize!.Value.ToPoints());
        Assert.False(runs[1].Bold);
        Assert.False(runs[1].Italic);
        Assert.False(runs[1].Strikethrough);
        Assert.Equal(ExcelUnderlineStyle.None, runs[1].UnderlineStyle);
        Assert.Equal(ExcelVerticalTextAlignment.Superscript, runs[1].VerticalTextAlignment);
    }
}
