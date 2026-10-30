using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void SlideAppearanceSurvivesContinuationAndNativeReopening() {
        using var document = PowerPointPresentation.Create();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            PowerPointSlide -Document $deck -Title 'Service details' -Layout TitleAndContent -Background Accent2 -Transition Fade -PassThru {
                PowerPointParagraph { PowerPointText ('Retained detail. ' * 40) -Size 72 }
            }
            """);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.True(document.Slides.Count > 1);
        Assert.Equal(document.Slides.Count, result.Count);
        for (int i = 0; i < result.Count; i++) Assert.Same(document.Slides[i], result[i].BaseObject);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.All(saved.Slides, slide => {
            Assert.Equal(saved.GetThemeColor(PowerPointThemeColor.Accent2), slide.GetBackground().Color);
            Assert.Equal(PowerPointSlideTransition.Fade, slide.Transition);
        });
        Assert.Equal(40, saved.Slides.SelectMany(slide => slide.TextBoxes).Sum(box =>
            System.Text.RegularExpressions.Regex.Matches(box.Text, "Retained detail\\.").Count));
    }

    [Fact]
    public void BackgroundImageUsesALiteralPathAndRetainsItsBytes() {
        string image = Path.Combine(_root, "background[1].png");
        byte[] bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        File.WriteAllBytes(image, bytes);
        using var document = PowerPointPresentation.Create();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("image", image);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("PowerPointSlide -Document $deck -Title 'Literal image' -BackgroundImage $image -Transition Fade -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Same(document.Slides[0], Assert.Single(result).BaseObject);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(bytes, saved.Slides[0].GetBackground().ImageBytes);
        Assert.Equal(PowerPointSlideTransition.Fade, saved.Slides[0].Transition);
    }

    [Theory]
    [InlineData("-Background Red -BackgroundImage $image")]
    [InlineData("-BackgroundImage $image")]
    [InlineData("-Transition 9999")]
    public void InvalidAppearanceDoesNotRunTheContentBlockOrAppendASlide(string arguments) {
        using var document = PowerPointPresentation.Create();
        var original = document.AddSlide();
        original.AddTitle("Preserved");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("image", Path.Combine(_root, "missing[1].png"));
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("$global:entered = $false; PowerPointSlide -Document $deck -Title Rejected " + arguments + " { $global:entered = $true }");
        var failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
        Assert.Same(original, Assert.Single(document.Slides));
        Assert.Contains(original.TextBoxes, box => box.Text == "Preserved");
    }

    [Fact]
    public void AppearanceFailureAfterCollectionRollsBackAndAllowsTheNextSlide() {
        string image = Path.Combine(_root, "disappearing.png");
        File.WriteAllBytes(image, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
        using var document = PowerPointPresentation.Create();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("image", image);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            PowerPointSlide -Document $deck -Title Rejected -BackgroundImage $image -PassThru -ErrorAction Continue {
                PowerPointParagraph 'Rejected content'
                [System.IO.File]::Delete($image)
            }
            PowerPointSlide -Document $deck -Title Recovered -PassThru { PowerPointList 'Accepted content' }
            """);
        var result = shell.Invoke();
        Assert.True(shell.HadErrors);
        Assert.Same(Assert.Single(document.Slides), Assert.Single(result).BaseObject);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Contains(saved.Slides[0].TextBoxes, box => box.Text == "Accepted content");
        Assert.DoesNotContain(saved.Slides[0].TextBoxes, box => box.Text.Contains("Rejected", StringComparison.Ordinal));
    }
}
