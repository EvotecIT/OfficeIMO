using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void PipelineChartDataAndSampleChartPersistOnSelectedSlidesWithPhysicalAndGridBounds() {
        string path = Path.Combine(_root, "selected-charts.pptx");
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
                @([pscustomobject]@{ Region='EMEA'; Revenue=10 }, [pscustomobject]@{ Region='APAC'; Revenue=20 }) |
                    PowerPointChart -Slide 2 -Category Region -Value Revenue -Type Bar -Column 7 -Span 6 -Y 5cm -Height 3cm
                PowerPointChart -Slide 1 -Type Pie -X 2cm -Y 4cm -Width 10cm -Height 5cm
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(2, saved.Slides.Count);
        var sample = Assert.Single(saved.Slides[0].Charts);
        Assert.True(sample.TryGetOfficeSnapshot(out var sampleData));
        Assert.Equal(OfficeChartType.Pie, sampleData.ChartKind);
        Assert.Equal(new[] { "Q1", "Q2", "Q3" }, sampleData.Data.Categories);
        Assert.Equal(new[] { 12d, 18d, 14d }, Assert.Single(sampleData.Data.Series).Values);
        Assert.Equal(720000L, sample.Left.ToEmus());
        Assert.Equal(1440000L, sample.Top.ToEmus());
        Assert.Equal(3600000L, sample.Width.ToEmus());
        Assert.Equal(1800000L, sample.Height.ToEmus());
        var chart = Assert.Single(saved.Slides[1].Charts);
        Assert.True(chart.TryGetOfficeSnapshot(out var projected));
        Assert.Equal(OfficeChartType.Bar, projected.ChartKind);
        Assert.Equal(new[] { "EMEA", "APAC" }, projected.Data.Categories);
        Assert.Equal(new[] { 10d, 20d }, Assert.Single(projected.Data.Series).Values);
        Assert.Equal(1800000L, chart.Top.ToEmus());
        Assert.Equal(1080000L, chart.Height.ToEmus());
        Assert.True(chart.Left.ToPoints() > saved.SlideSize.Width.ToPoints() / 2);
        Assert.True(chart.Width.ToPoints() > 0 && chart.Width.ToPoints() < saved.SlideSize.Width.ToPoints() / 2);
        Assert.True(chart.Left + chart.Width < saved.SlideSize.Width);
    }
}
