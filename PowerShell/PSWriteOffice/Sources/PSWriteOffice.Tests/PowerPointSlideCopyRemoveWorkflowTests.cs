using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointSlideEditWorkflowTests {
    [Fact]
    public void CopyWithoutAPositionInsertsAfterTheSourceAndIsQuietByDefault() {
        using var document = Create();
        PowerPointSlide second = document.Slides[1];
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Copy-OfficePowerPointSlide -Document $deck -Slide 1");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(3, document.Slides.Count);
        Assert.Same(second, document.Slides[2]);
        Assert.Equal("First", document.Slides[1].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
    }

    [Theory]
    [InlineData("Copy-OfficePowerPointSlide -Document $deck -Slide 1")]
    [InlineData("Copy-OfficePowerPointSlide -Slide $deck.Slides[0]")]
    [InlineData("$deck.Slides[0] | Copy-OfficePowerPointSlide")]
    public void CopyUsesOneBasedSelectionAndPreservesNativeNotes(string command) {
        using var document = Create();
        document.Slides[0].Notes.Text = "Retained source notes";
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript(command + " -Position 3 -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(3, document.Slides.Count);
        Assert.Same(document.Slides[2], Assert.Single(result).BaseObject);
        Assert.NotSame(document.Slides[0], document.Slides[2]);
        document.Slides[2].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text = "Copied appendix";
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal("Retained source notes", saved.Slides[2].Notes.Text);
        Assert.Equal("First", saved.Slides[0].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
        Assert.Equal("Copied appendix", saved.Slides[2].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
    }

    [Fact]
    public void RemovePassThruReturnsTheLiveOwningDocumentForFurtherEditing() {
        using var document = Create();
        document.AddSlide().AddTitle("Third");
        PowerPointSlide survivor = document.Slides[1];
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Remove-OfficePowerPointSlide -Document $deck -Slide 1, 3 -PassThru -Confirm:$false");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Same(document, Assert.Single(result).BaseObject);
        Assert.Same(survivor, Assert.Single(document.Slides));
        document.AddSlide().AddTitle("Replacement");
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(2, saved.Slides.Count);
    }

    [Theory]
    [InlineData("Remove-OfficePowerPointSlide -Document $deck -Slide 1, 3")]
    [InlineData("Remove-OfficePowerPointSlide -Slide $deck.Slides[0], $deck.Slides[2]")]
    [InlineData("@($deck.Slides[0], $deck.Slides[2]) | Remove-OfficePowerPointSlide")]
    public void RemoveKeepsTheUnselectedSlideAcrossArraysAndPipelines(string command) {
        using var document = Create();
        document.AddSlide().AddTitle("Third");
        PowerPointSlide survivor = document.Slides[1];
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript(command + " -Confirm:$false");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Same(survivor, Assert.Single(document.Slides));
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal("Second", saved.Slides[0].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
    }

    [Theory]
    [InlineData("Remove-OfficePowerPointSlide -Document $deck -Slide 1, 99 -Confirm:$false")]
    [InlineData("Copy-OfficePowerPointSlide -Document $deck -Slide 1 -Position 99")]
    [InlineData("Copy-OfficePowerPointSlide -Document $other -Slide $deck.Slides[0]")]
    [InlineData("Remove-OfficePowerPointSlide -Document $other -Slide $deck.Slides[0] -Confirm:$false")]
    public void InvalidCopyOrRemovalLeavesBothDocumentsUnchanged(string command) {
        using var document = Create(); using var other = Create();
        PowerPointSlide first = document.Slides[0], second = document.Slides[1];
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("other", other);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript(command);
        var failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.DoesNotContain(shell.Streams.Error, error => error.Exception is CommandNotFoundException);
        Assert.Equal(2, document.Slides.Count);
        Assert.Equal(2, other.Slides.Count);
        Assert.Same(first, document.Slides[0]); Assert.Same(second, document.Slides[1]);
    }

    [Fact]
    public void RemoveWhatIfKeepsTheNativeSlidesAndProducesNoDocumentOutput() {
        using var document = Create();
        PowerPointSlide first = document.Slides[0], second = document.Slides[1];
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Remove-OfficePowerPointSlide -Document $deck -Slide 1, 2 -PassThru -WhatIf");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(2, document.Slides.Count);
        Assert.Same(first, document.Slides[0]); Assert.Same(second, document.Slides[1]);
    }
}
