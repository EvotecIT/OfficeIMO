using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed partial class PowerPointSlideEditWorkflowTests {
    [Theory]
    [InlineData("Set-OfficePowerPointSlide 1 -Document $deck")]
    [InlineData("Set-OfficePowerPointSlide -Slide $deck.Slides[0]")]
    [InlineData("$deck.Slides[0] | Set-OfficePowerPointSlide")]
    public void SlideSelectionEditsAndReturnsTheOriginalNativeObject(string command) {
        using var document = Create();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript(command + " -Title 'Updated' -Background Accent2 -Transition Fade -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Same(document.Slides[0], Assert.Single(result).BaseObject);
        Assert.Equal("Updated", document.Slides[0].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
        Assert.Equal(PowerPointSlideTransition.Fade, document.Slides[0].Transition);
        Assert.Equal(document.GetThemeColor(PowerPointThemeColor.Accent2), document.Slides[0].GetBackground().Color);
        Assert.Equal("Second", document.Slides[1].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
    }

    [Theory]
    [InlineData("Set-OfficePowerPointSlide 0 -Document $deck")]
    [InlineData("Set-OfficePowerPointSlide 3 -Document $deck")]
    [InlineData("Set-OfficePowerPointSlide 1 -Document $deck -Master 2")]
    [InlineData("Set-OfficePowerPointSlide 1 -Document $deck -TemplateLayout 'Missing layout'")]
    [InlineData("Set-OfficePowerPointSlide -Document $other -Slide $deck.Slides[0]")]
    public void InvalidSelectorsLeaveTheOriginalTitleAndBackgroundUnchanged(string command) {
        using var document = Create(); using var other = Create();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document); runspace.SessionStateProxy.SetVariable("other", other);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript(command + " -Title 'Must not apply' -Background Red -PassThru");
        Exception? failure = null;
        var result = new System.Collections.ObjectModel.Collection<PSObject>();
        try { result = shell.Invoke(); } catch (Exception error) { failure = error; }
        Assert.True(failure != null || shell.HadErrors);
        Assert.Empty(result);
        Assert.Equal("First", document.Slides[0].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
        Assert.Null(document.Slides[0].BackgroundColor);
    }

    [Fact]
    public void TemplateMetadataAndBackgroundClearingUseTheSelectedSlide() {
        using var document = Create();
        var layout = PowerPointTemplate.Inspect(document).Masters[0].Layouts.Last();
        document.Slides[0].SetBackgroundGradient("112233", "445566");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document); runspace.SessionStateProxy.SetVariable("layout", layout);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Set-OfficePowerPointSlide 1 -Document $deck -TemplateLayout $layout -ClearBackground");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(layout.LayoutIndex, document.Slides[0].LayoutIndex);
        Assert.Equal(document.Slides[1].GetBackground().Color, document.Slides[0].GetBackground().Color);
    }

    [Fact]
    public void LiteralImagePathsAndTheOriginalSlideSurviveNativeSaveAndReopen() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-slide-edit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string image = Path.Combine(root, "background[1].png");
            byte[] bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
            File.WriteAllBytes(image, bytes);
            using var document = Create();
            using var runspace = Open();
            runspace.SessionStateProxy.SetVariable("deck", document); runspace.SessionStateProxy.SetVariable("image", image);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("Set-OfficePowerPointSlide 1 -Document $deck -BackgroundImage $image -PassThru");
            var result = shell.Invoke();
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Same(document.Slides[0], Assert.Single(result).BaseObject);
            using var input = new MemoryStream(document.ToBytes());
            using var reopened = PowerPointPresentation.Load(input);
            Assert.Equal(PowerPointSlideBackgroundKind.Image, reopened.Slides[0].GetBackground().Kind);
            Assert.Equal(bytes, reopened.Slides[0].GetBackground().ImageBytes);
            Assert.NotEqual(PowerPointSlideBackgroundKind.Image, reopened.Slides[1].GetBackground().Kind);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void ClearingTheTitlePreservesUnboundBackgroundAndTransition() {
        using var document = Create();
        document.Slides[0].SetBackgroundColor(OfficeIMO.OfficeColor.FromRgba(17, 34, 51, 128));
        document.Slides[0].Transition = PowerPointSlideTransition.Fade;
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Set-OfficePowerPointSlide 1 -Document $deck -Title ''");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(string.Empty, document.Slides[0].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
        Assert.Equal("11223380", document.Slides[0].GetBackground().Color);
        Assert.Equal(PowerPointSlideTransition.Fade, document.Slides[0].Transition);
    }

    private static PowerPointPresentation Create() {
        var document = PowerPointPresentation.Create(options: new PowerPointCreateOptions { PersistenceMode = OfficeIMO.DocumentPersistenceMode.Explicit });
        document.AddSlide().AddTitle("First"); document.AddSlide().AddTitle("Second");
        return document;
    }
    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Set-OfficePowerPointSlide", typeof(SetOfficePowerPointSlideCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Copy-OfficePowerPointSlide", typeof(CopyOfficePowerPointSlideCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Remove-OfficePowerPointSlide", typeof(RemoveOfficePowerPointSlideCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open(); return runspace;
    }
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString()));
}
