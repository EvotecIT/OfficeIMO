using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData("-InputObject @([pscustomobject]@{ Product = 'Alpha'; Qty = 12 })", true)]
    [InlineData("-Rows 2 -Columns 2", false)]
    public void ExplicitTableBoundsInsideSlideBlocksSurviveLayoutAndDoNotOverlapFixedFigures(string data, bool headers) {
        string path = Path.Combine(_root, "manual-table.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript($$"""
            New-OfficePowerPoint -Path $target {
                PowerPointSlide -Title 'Manual table' {
                    PowerPointTable {{data}} -X 60 -Y 140 -Width 420 -Height 200
                    PowerPointShape -Type Ellipse -Name Signal -X 520 -Y 140 -Width 100 -Height 100
                }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var slide = Assert.Single(saved.Slides);
        var table = Assert.Single(slide.Tables);
        Assert.Equal(762000L, table.Left.ToEmus());
        Assert.Equal(1778000L, table.Top.ToEmus());
        Assert.Equal(5334000L, table.Width.ToEmus());
        Assert.Equal(2540000L, table.Height.ToEmus());
        Assert.Equal(2, table.Rows);
        if (headers) {
            Assert.Equal("Product", table.GetCell(0, 0).Text);
            Assert.Equal("Alpha", table.GetCell(1, 0).Text);
        }
        var signal = Assert.Single(slide.Shapes, shape => shape.Name == "Signal");
        Assert.True(table.Left + table.Width <= signal.Left);
    }
}
