using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SectionBoundariesIncludeFollowingSlidesAndTheirContinuations(bool cover) {
        string path = Path.Combine(_root, "sections.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        runspace.SessionStateProxy.SetVariable("cover", cover);
        runspace.SessionStateProxy.SetVariable("body", string.Join("\n", Enumerable.Range(1, 80).Select(i => "Regional sales review " + i)));
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target {
                if ($cover) { PowerPointSlide -Title 'Overview' }
                $script:results = PowerPointSection 'Results' -PassThru
                PowerPointSlide -Title 'Revenue' -Layout TitleAndContent {
                    PowerPointParagraph $body -Size 18
                }
                PowerPointSlide -Title 'Margin' { PowerPointParagraph 'Margin improved.' }
                PowerPointSection 'Plans'
                PowerPointSlide -Title 'Priorities' { PowerPointParagraph 'Review capacity.' }
            }
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        object output = runspace.SessionStateProxy.GetVariable("results");
        var returned = Assert.IsType<PowerPointSectionInfo>(output is PSObject wrapped ? wrapped.BaseObject : output);
        using var saved = PowerPointPresentation.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.True(saved.Slides.Count > (cover ? 4 : 3), "The long paragraph must exercise physical continuation slides.");
        var results = Assert.Single(saved.GetSections(), section => section.Name == "Results");
        var plans = Assert.Single(saved.GetSections(), section => section.Name == "Plans");
        Assert.Equal(returned.Id, results.Id);
        int first = cover ? 1 : 0;
        int plansStart = Assert.Single(plans.SlideIndices);
        Assert.Equal(Enumerable.Range(first, plansStart - first), results.SlideIndices);
        if (cover) Assert.Equal(new[] { 0 }, saved.GetSections()[0].SlideIndices);
    }

    [Theory]
    [InlineData("PowerPointSection 'Results' -Document $deck -Slide 2")]
    [InlineData("PowerPointSection 'Results' -Slide $deck.Slides[1]")]
    [InlineData("$deck | Add-OfficePowerPointSection 'Results' -Slide 2")]
    public void SectionsSelectExistingSlidesByOneBasedPositionOrNativeOwnership(string command) {
        using var document = PowerPointPresentation.Create();
        document.AddSlide(); document.AddSlide(); document.AddSlide();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript(command + " -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var section = Assert.IsType<PowerPointSectionInfo>(Assert.Single(result).BaseObject);
        Assert.Equal(new[] { 1, 2 }, section.SlideIndices);
        Assert.Equal(section.Id, Assert.Single(document.GetSections(), item => item.Name == "Results").Id);
    }

    [Fact]
    public void ARejectedSlideCannotLeaveItsSectionOrAnUnrequestedDefaultSection() {
        using var document = PowerPointPresentation.Create();
        document.AddSlide().AddTitle("Keep");
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            try {
                PowerPointSlide -Document $deck -Title 'Rejected' {
                    PowerPointSection 'Rejected section'
                    throw 'Reject this slide'
                } -ErrorAction Stop
            } catch { $_.Exception.Message }
            """);
        Assert.Contains("Reject this slide", Assert.IsType<string>(Assert.Single(shell.Invoke()).BaseObject));
        using var accepted = PowerShell.Create(); accepted.Runspace = runspace;
        accepted.AddScript("PowerPointSlide -Document $deck -Title 'Accepted' { PowerPointParagraph 'Keep this content.' }");
        Assert.Empty(accepted.Invoke());
        Assert.False(accepted.HadErrors, Errors(accepted));
        Assert.Equal(2, document.Slides.Count);
        Assert.Empty(document.GetSections());
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Empty(saved.GetSections());
    }

    [Fact]
    public void ARejectedColumnCannotLeaveItsSectionOnTheAcceptedSlide() {
        using var document = PowerPointPresentation.Create();
        document.AddSlide().AddTitle("Keep");
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            PowerPointSlide -Document $deck -Title 'Accepted' {
                try {
                    PowerPointColumn {
                        PowerPointSection 'Rejected section'
                        throw 'Reject this column'
                    } -ErrorAction Stop
                } catch { }
                PowerPointParagraph 'Retained content.'
            }
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(2, saved.Slides.Count);
        Assert.Empty(saved.GetSections());
    }

    [Fact]
    public void ARejectedFirstSlideLeavesThePendingSectionForTheNextAcceptedSlide() {
        using var document = PowerPointPresentation.Create();
        document.AddSlide().AddTitle("Overview");
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            PowerPointSection 'Results' -Document $deck
            try {
                PowerPointSlide -Document $deck -Title 'Rejected' { throw 'Reject this slide' } -ErrorAction Stop
            } catch { $_.Exception.Message }
            """);
        Assert.Contains("Reject this slide", Assert.IsType<string>(Assert.Single(shell.Invoke()).BaseObject));
        using var accepted = PowerShell.Create(); accepted.Runspace = runspace;
        accepted.AddScript("""
            PowerPointSlide -Document $deck -Title 'Accepted' { PowerPointParagraph 'Retained.' }
            PowerPointSlide -Document $deck -Title 'Following' { PowerPointParagraph 'Also retained.' }
            """);
        Assert.Empty(accepted.Invoke());
        Assert.False(accepted.HadErrors, Errors(accepted));
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(3, saved.Slides.Count);
        Assert.Equal(new[] { 1, 2 }, Assert.Single(saved.GetSections(), section => section.Name == "Results").SlideIndices);
        Assert.Equal(new[] { 0 }, saved.GetSections()[0].SlideIndices);
    }

    [Theory]
    [InlineData("PowerPointSection ' ' -Document $deck -Slide 1")]
    [InlineData("PowerPointSection 'Rejected' -Document $deck -Slide 0")]
    [InlineData("PowerPointSection 'Rejected' -Document $other -Slide $deck.Slides[0]")]
    public void RejectedSectionInputsCannotModifyEitherDeck(string command) {
        using var document = PowerPointPresentation.Create(); document.AddSlide();
        using var other = PowerPointPresentation.Create(); other.AddSlide();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("other", other);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript(command);
        Exception? failure = null;
        try { Assert.Empty(shell.Invoke()); } catch (RuntimeException error) { failure = error; }
        Assert.True(failure != null || shell.HadErrors);
        Assert.Empty(document.GetSections());
        Assert.Empty(other.GetSections());
        Assert.Single(document.Slides); Assert.Single(other.Slides);
    }
}
