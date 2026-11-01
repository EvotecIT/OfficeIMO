using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void ManualTextContentUsesTheSelectedNativeBoxAndPersistsItsRunsAndBounds() {
        string path = Path.Combine(_root, "manual-text.pptx");
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("slide", slide);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $slide | PowerPointTextBox -Name 'Manual status' -X 2cm -Y 3cm -Width 12cm -Height 4cm -PassThru {
                PowerPointParagraph {
                    PowerPointText 'Ready ' -Bold -Color '#112233'
                    PowerPointText 'now' -Italic
                }
                PowerPointParagraph 'Next review tomorrow'
            }
            """);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var box = Assert.IsType<PowerPointTextBox>(Assert.Single(result).BaseObject);
        Assert.Same(box, Assert.Single(slide.TextBoxes));
        Assert.Single(document.Slides);
        document.Save(path);
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var persisted = Assert.Single(Assert.Single(saved.Slides).TextBoxes);
        Assert.Equal("Manual status", persisted.Name);
        Assert.Equal(2, persisted.Paragraphs.Count);
        Assert.Equal("Ready now", persisted.Paragraphs[0].Text);
        Assert.Equal("Next review tomorrow", persisted.Paragraphs[1].Text);
        Assert.Equal(2, persisted.Paragraphs[0].Runs.Count);
        Assert.True(persisted.Paragraphs[0].Runs[0].Bold);
        Assert.True(persisted.Paragraphs[0].Runs[1].Italic);
        Assert.Equal("112233", persisted.Paragraphs[0].Runs[0].Color);
        Assert.Equal(720000L, persisted.Left.ToEmus());
        Assert.Equal(1080000L, persisted.Top.ToEmus());
        Assert.Equal(4320000L, persisted.Width.ToEmus());
        Assert.Equal(1440000L, persisted.Height.ToEmus());
    }

    [Theory]
    [InlineData("PowerPointTextBox")]
    [InlineData("PowerPointParagraph")]
    [InlineData("PowerPointList")]
    public void FailedManualTextContentRemovesItsBoxAndRestoresTheTextContext(string command) {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var existing = slide.AddTextBox("Keep this");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("slide", slide);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        string content = command == "PowerPointList" ? "PowerPointListItem { PowerPointText 'Partial' }" : "PowerPointText 'Partial'";
        shell.AddScript("$slide | " + command + " -X 1cm -Y 2cm -Width 3cm -Height 1cm -ErrorAction Stop { " + content + "; throw 'manual-content-failed' }");
        Exception? error = Record.Exception(() => shell.Invoke());
        Assert.True(error != null || shell.HadErrors);
        Assert.Contains("manual-content-failed", error?.ToString() + Errors(shell));
        Assert.Same(existing, Assert.Single(slide.TextBoxes));
        shell.Commands.Clear();
        shell.Streams.Error.Clear();
        shell.AddScript("$slide | PowerPointTextBox { PowerPointText 'Recovered' } -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal("Recovered", Assert.IsType<PowerPointTextBox>(Assert.Single(result).BaseObject).Text);
        Assert.Equal(2, slide.TextBoxes.Count());
        Assert.Single(document.Slides);
    }

    [Fact]
    public void ManualGridBoxesRemainFixedWhenTheSlideAlsoContainsFlowingContent() {
        string path = Path.Combine(_root, "manual-grid.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target -Theme Modern {
                PowerPointSlide -Title 'Manual and flowing content' {
                    PowerPointTextBox 'Left' -Column 1 -Span 6 -Y 5cm -Height 1cm -Name Left
                    PowerPointTextBox 'Right' -Column 7 -Span 6 -Y 5cm -Height 1cm -Name Right
                    PowerPointShape -Type Ellipse -Name Signal -Column 1 -Span 6 -X 2cm -Width 2cm -Y 7cm -Height 2cm -FillColor '#112233' -OutlineColor '#445566' -OutlineWidth 0.5mm
                    PowerPointCard 'Flowing' -Text 'Semantic content remains editable'
                }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var slide = Assert.Single(saved.Slides);
        var left = Assert.Single(slide.TextBoxes, box => box.Name == "Left");
        var right = Assert.Single(slide.TextBoxes, box => box.Name == "Right");
        Assert.Equal("Left", left.Text);
        Assert.Equal("Right", right.Text);
        Assert.Equal(1800000L, left.Top.ToEmus());
        Assert.Equal(left.Top, right.Top);
        Assert.True(left.Width.ToEmus() > 0);
        Assert.Equal(left.Width, right.Width);
        Assert.True(left.Left.ToEmus() >= 0);
        Assert.True(left.Left.ToEmus() + left.Width.ToEmus() <= right.Left.ToEmus());
        Assert.Contains(slide.TextBoxes, box => box.Text.Contains("Semantic content remains editable"));
        var signal = Assert.IsType<PowerPointAutoShape>(Assert.Single(slide.Shapes, shape => shape.Name == "Signal"));
        Assert.Equal(OfficeIMO.Drawing.OfficePresetShapeType.Ellipse, signal.ShapeType);
        Assert.Equal("112233", signal.FillColor, ignoreCase: true);
        Assert.Equal("445566", signal.OutlineColor, ignoreCase: true);
        Assert.Equal(720000L, signal.Left.ToEmus());
        Assert.Equal(720000L, signal.Width.ToEmus());
        Assert.Equal(2520000L, signal.Top.ToEmus());
        Assert.Equal(18000L, signal.OutlineWidth!.Value.ToEmus());
    }

    [Fact]
    public void PhysicalLayoutBoxesCanBeUsedDirectlyInAManualTextCommand() {
        using var document = PowerPointDocument.Create();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $columns = Get-OfficePowerPointLayoutBox -Document $deck -Columns 2 -Margin 1.5cm -Gutter 1cm
            PowerPointSlide -Document $deck -Title 'Status' {
                PowerPointTextBox 'Healthy' -X $columns[0].Left -Y $columns[0].Top -Width $columns[0].Width -Height 40
                PowerPointTextBox 'Full grid' -Column 1 -Y 8cm -Name FullGrid
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var slide = Assert.Single(document.Slides);
        var label = Assert.Single(slide.TextBoxes, box => box.Text == "Healthy");
        Assert.Equal(540000L, label.Left.ToEmus());
        Assert.Equal(540000L, label.Top.ToEmus());
        Assert.Equal(508000L, label.Height.ToEmus());
        Assert.True(label.Width.ToEmus() > 0);
        var grid = Assert.Single(slide.TextBoxes, box => box.Name == "FullGrid");
        Assert.True(grid.Width > label.Width);
    }
}
