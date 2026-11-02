using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData("@([pscustomobject]@{ Amount = [decimal]1.5 })", false)]
    [InlineData("@( ,@('Amount'); ,@([decimal]1.5) )", false)]
    [InlineData("@( ,@('Amount'); ,@(@{ Text = [decimal]1.5; Bold = $true }) )", true)]
    public void TableInputsPersistPhysicalLengthsNamedStyleAndLocalizedData(string data, bool rich) {
        string path = Path.Combine(_root, "localized-table.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript($$"""
            New-OfficePowerPoint -Path $target {
                PowerPointSlide -Title 'Localized data' {
                    PowerPointTable -InputObject {{data}} -Culture fr-FR -Style 'Medium Style 2 - Accent 2' -X 2.5cm -Y 5cm -Width 12cm -Height 3cm
                }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var table = Assert.Single(Assert.Single(saved.Slides).Tables);
        Assert.Equal(2, table.Rows);
        Assert.Equal("Amount", table.GetCell(0, 0).Text);
        var amount = table.GetCell(1, 0);
        Assert.Equal("1,5", amount.Text);
        if (rich) Assert.True(Assert.Single(amount.Runs).Bold);
        Assert.Equal("{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}", table.StyleId, ignoreCase: true);
        Assert.Equal(900000L, table.Left.ToEmus());
        Assert.Equal(1800000L, table.Top.ToEmus());
        Assert.Equal(4320000L, table.Width.ToEmus());
        Assert.Equal(1080000L, table.Height.ToEmus());
    }

    [Fact]
    public void TableNumericSelectorAndGridEditTheSecondSlideOfAnOpenBlock() {
        string path = Path.Combine(_root, "selected-table.pptx");
        using (var document = PowerPointDocument.Create()) {
            document.AddSlide();
            document.AddSlide();
            document.Save(path);
        }
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            Open-OfficePowerPoint -Path $target {
                PowerPointTable -Slide 2 -Rows 2 -Columns 1 -Column 7 -Span 6 -Y 5cm -Height 3cm -Style '{5940675A-B579-460E-94D1-54222C63F5DA}'
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(2, saved.Slides.Count);
        Assert.Empty(saved.Slides[0].Tables);
        var table = Assert.Single(saved.Slides[1].Tables);
        Assert.Equal(2, table.Rows);
        Assert.Equal("{5940675A-B579-460E-94D1-54222C63F5DA}", table.StyleId, ignoreCase: true);
        Assert.Equal(1800000L, table.Top.ToEmus());
        Assert.True(table.Left.ToPoints() > saved.SlideSize.Width.ToPoints() / 2);
        Assert.True(table.Width.ToPoints() > 0 && table.Width.ToPoints() < saved.SlideSize.Width.ToPoints() / 2);
        Assert.True(table.Left + table.Width < saved.SlideSize.Width);
    }
}
