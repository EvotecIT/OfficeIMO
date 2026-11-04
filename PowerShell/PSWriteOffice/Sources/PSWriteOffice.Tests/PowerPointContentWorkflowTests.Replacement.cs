using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void ReplacementUsesDocumentPipelineAndSlideNumberWhileNoTablesPreservesCells() {
        string path = Path.Combine(_root, "selected-replacement.pptx");
        using (var document = PowerPointDocument.Create()) {
            for (int index = 0; index < 2; index++) {
                var slide = document.AddSlide();
                slide.AddTextBox("Old label");
                slide.AddTable(1, 1).GetCell(0, 0).Text = "Old data";
            }
            document.Save(path);
        }
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $deck = Open-OfficePowerPoint -Path $target
            try {
                $global:slideCount = $deck | Update-OfficePowerPointText 'Old' 'New' -Slide 2 -NoTables -PassThru
                $global:deckCount = $deck | Update-OfficePowerPointText 'Old' 'Archived' -PassThru
                Save-OfficePowerPoint -Document $deck
            } finally { Close-OfficePowerPoint -Document $deck }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(1, Assert.IsType<int>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("slideCount")).BaseObject));
        Assert.Equal(3, Assert.IsType<int>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("deckCount")).BaseObject));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal("Archived label", Assert.Single(saved.Slides[0].TextBoxes).Text);
        Assert.Equal("New label", Assert.Single(saved.Slides[1].TextBoxes).Text);
        foreach (var slide in saved.Slides) Assert.Equal("Archived data", Assert.Single(slide.Tables).GetCell(0, 0).Text);
    }

    [Fact]
    public void ReplacementInsideContentBlocksUsesTheCurrentSlideThenTheOpenDocument() {
        string path = Path.Combine(_root, "scoped-replacement.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target {
                PowerPointSlide -Title 'First' { PowerPointParagraph 'Old first' }
                PowerPointSlide -Title 'Second' {
                    PowerPointParagraph 'Old second'
                    $global:slideCount = Update-OfficePowerPointText 'Old' 'New' -PassThru
                }
            }
            Open-OfficePowerPoint -Path $target {
                $global:deckCount = Update-OfficePowerPointText 'Old' 'Final' -PassThru
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(1, Assert.IsType<int>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("slideCount")).BaseObject));
        Assert.Equal(1, Assert.IsType<int>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("deckCount")).BaseObject));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Contains(saved.Slides[0].TextBoxes, box => box.Text == "Final first");
        Assert.Contains(saved.Slides[1].TextBoxes, box => box.Text == "New second");
        Assert.DoesNotContain(saved.Slides[0].TextBoxes, box => box.Text.Contains("New first"));
    }
}
