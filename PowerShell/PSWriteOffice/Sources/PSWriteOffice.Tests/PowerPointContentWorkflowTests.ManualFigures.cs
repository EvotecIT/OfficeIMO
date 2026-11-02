using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExplicitImageAndChartBoundsInsideSlideBlocksRemainFixed(bool image) {
        string path = Path.Combine(_root, "manual-figure.pptx");
        string png = Path.Combine(_root, "marker.png");
        var pixels = new byte[5 * 3 * 4];
        for (int index = 0; index < pixels.Length; index += 4) {
            pixels[index + 1] = 100;
            pixels[index + 2] = 180;
            pixels[index + 3] = 255;
        }
        File.WriteAllBytes(png, OfficePngWriter.EncodeRgba(5, 3, pixels));
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        runspace.SessionStateProxy.SetVariable("png", png);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        string figure = image ? "PowerPointImage -Path $png -AlternativeText 'Status marker'" :
            "PowerPointChart -InputObject @([pscustomobject]@{ Quarter='Q1'; Revenue=10 }, [pscustomobject]@{ Quarter='Q2'; Revenue=20 }) -Category Quarter -Value Revenue -Type Column";
        shell.AddScript($$"""
            New-OfficePowerPoint -Path $target {
                PowerPointSlide -Title 'Manual figure' {
                    {{figure}} {{(image ? "-X 60 -Y 140 -Width 200 -Height 120" : "-X 60pt -Y 140pt -Width 200pt -Height 120pt")}}
                    PowerPointShape -Type Rectangle -Name Neighbor -X 300 -Y 140 -Width 80 -Height 80
                }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var slide = Assert.Single(saved.Slides);
        PowerPointShape persisted = image ? Assert.Single(slide.Pictures) : Assert.Single(slide.Charts);
        Assert.Equal(762000L, persisted.Left.ToEmus());
        Assert.Equal(1778000L, persisted.Top.ToEmus());
        Assert.Equal(2540000L, persisted.Width.ToEmus());
        Assert.Equal(1524000L, persisted.Height.ToEmus());
        var neighbor = Assert.Single(slide.Shapes, shape => shape.Name == "Neighbor");
        Assert.True(persisted.Left + persisted.Width <= neighbor.Left);
        if (image) Assert.Equal("Status marker", persisted.Description);
    }
}
