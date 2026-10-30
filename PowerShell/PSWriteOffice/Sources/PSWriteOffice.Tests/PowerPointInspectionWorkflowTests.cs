using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PowerPointInspectionWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-ppt-inspection-" + Guid.NewGuid().ToString("N"));
    public PowerPointInspectionWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public void LiveSlideSelectionUsesOneBasedPositionsAndReturnsTheOriginalObjects() {
        using var document = Create();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("$deck | Get-OfficePowerPointSlide -Slide 2,1");
        var results = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(2, results.Count);
        Assert.Same(document.Slides[1], results[0].BaseObject);
        Assert.Same(document.Slides[0], results[1].BaseObject);
    }

    [Fact]
    public void InspectionUsesTheEnclosingDocumentAndDoesNotCreateNotes() {
        string target = Path.Combine(_root, "scope.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", target);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target {
                Add-OfficePowerPointSlide -Title 'Review'
                $global:slides = @(Get-OfficePowerPointSlide -Summary)
                $global:notes = @(Get-OfficePowerPointNotes -IncludeEmpty)
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var summary = Assert.IsType<PowerPointSlideSummaryInfo>(Assert.Single((object[])runspace.SessionStateProxy.GetVariable("slides")) is PSObject s ? s.BaseObject : null);
        Assert.Equal("Review", summary.Title);
        Assert.False(summary.HasNotes);
        using var saved = PowerPointPresentation.Load(target);
        Assert.False(Assert.Single(saved.Slides).HasSpeakerNotes);
    }

    [Fact]
    public void LiteralFilePipelineReturnsDetachedSummariesAndReleasesTheFile() {
        string path = Path.Combine(_root, "review[1].pptx");
        using (var document = Create()) { document.Slides[1].Notes.Text = "Explain the margin."; document.Save(path); }
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("file", new FileInfo(path));
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("$file | Get-OfficePowerPointSlide -Summary -Slide 2");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var summary = Assert.IsType<PowerPointSlideSummaryInfo>(Assert.Single(result).BaseObject);
        Assert.Equal("Second", summary.Title);
        Assert.True(summary.HasNotes);
        Assert.Equal("Explain the margin.", summary.NotesText);
        Assert.Equal(2, summary.SlideIndex);
        Assert.Null(summary.Slide);
        Assert.Equal(path, summary.Path);
        using var exclusive = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.True(exclusive.Length > 0);
    }

    [Fact]
    public void WildcardFileSummariesKeepValidItemsWhenAnotherInputIsMissing() {
        string path = Path.Combine(_root, "review.pptx");
        using (var document = Create()) document.Save(path);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("pattern", Path.Combine(_root, "*.pptx"));
        runspace.SessionStateProxy.SetVariable("missing", Path.Combine(_root, "missing.pptx"));
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Get-OfficePowerPointSlide -Path $missing,$pattern -Summary -Slide 1 -ErrorAction Continue");
        var result = shell.Invoke();
        Assert.True(shell.HadErrors);
        var summary = Assert.IsType<PowerPointSlideSummaryInfo>(Assert.Single(result).BaseObject);
        Assert.Equal("First", summary.Title);
        Assert.Equal(path, summary.Path);
    }

    [Theory]
    [InlineData("Get-OfficePowerPointSlide -Path $path")]
    [InlineData("Get-OfficePowerPointSlide -Document $deck -Slide 0")]
    [InlineData("Get-OfficePowerPointSlide -Document $deck -Slide 3")]
    [InlineData("Get-OfficePowerPointNotes -Document $deck -Slide 0")]
    public void InvalidSelectionsDoNotReturnOtherSlides(string script) {
        using var document = Create();
        string path = Path.Combine(_root, "review.pptx"); document.Save(path);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("path", path);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript(script);
        Exception? failure = null;
        var result = new System.Collections.ObjectModel.Collection<PSObject>();
        try { result = shell.Invoke(); } catch (Exception error) { failure = error; }
        Assert.True(failure != null || shell.HadErrors);
        Assert.Empty(result);
        Assert.Equal(2, document.Slides.Count);
    }

    [Theory]
    [InlineData("1")]
    [InlineData("$deck.Slides[0]")]
    public void NotesSelectorsReadExistingTextWithoutCreatingEmptyNotes(string selector) {
        using var document = Create(); document.Slides[0].Notes.Text = "Explain the decision.";
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Get-OfficePowerPointNotes -Document $deck -Slide " + selector);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var notes = Assert.IsType<PowerPointNotesInfo>(Assert.Single(result).BaseObject);
        Assert.Equal("Explain the decision.", notes.Text);
        Assert.Equal(1, notes.SlideIndex);
        Assert.Same(document.Slides[0], notes.Slide);
        Assert.False(document.Slides[1].HasSpeakerNotes);
    }

    [Fact]
    public void SummaryTitlesKeepTheWordBoundaryOfPhysicalLineBreaks() {
        using var document = Create();
        var title = document.Slides[0].TextBoxes.First();
        title.Text = "Customer";
        title.Paragraphs[0].AddLineBreak().AddText("onboarding review");
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Get-OfficePowerPointSlide -Document $deck -Slide 1 -Summary");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal("Customer onboarding review", Assert.IsType<PowerPointSlideSummaryInfo>(Assert.Single(result).BaseObject).Title);
        Assert.Equal("Customer" + Environment.NewLine + "onboarding review", title.Text);
    }

    private static PowerPointPresentation Create() {
        var document = PowerPointPresentation.Create(options: new PowerPointCreateOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
        document.AddSlide().AddTextBox("First"); document.AddSlide().AddTextBox("Second");
        return document;
    }
    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        foreach (var type in new[] { typeof(NewOfficePowerPointCommand), typeof(AddOfficePowerPointSlideCommand), typeof(GetOfficePowerPointSlideCommand), typeof(GetOfficePowerPointNotesCommand) }) {
            var cmdlet = type.GetCustomAttributes(typeof(CmdletAttribute), false).Cast<CmdletAttribute>().Single();
            state.Commands.Add(new SessionStateCmdletEntry(cmdlet.VerbName + "-" + cmdlet.NounName, type, null));
        }
        var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open(); return runspace;
    }
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString()));
}
