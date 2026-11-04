using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void AnUnusedEarlierSectionCannotCaptureLaterSlidesOrTheirContinuations(bool cover, bool reopenBeforeAppend) {
        using var document = PowerPointDocument.Create();
        if (cover) document.AddSlide().AddTitle("Overview");
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var start = PowerShell.Create(); start.Runspace = runspace;
        start.AddScript("""
            PowerPointSection Optional -Document $deck -PassThru
            PowerPointSection Results -Document $deck -PassThru
            """);
        var boundaries = start.Invoke().Select(item => Assert.IsType<PowerPointSectionInfo>(item.BaseObject)).ToArray();
        Assert.False(start.HadErrors, Errors(start)); Assert.Equal(2, boundaries.Length);
        using var reopened = reopenBeforeAppend ? PowerPointDocument.Load(new MemoryStream(document.ToBytes())) : null;
        PowerPointDocument active = reopened ?? document;
        runspace.SessionStateProxy.SetVariable("deck", active);
        runspace.SessionStateProxy.SetVariable("body", string.Join("\n", Enumerable.Range(1, 80).Select(i => "Regional results " + i)));
        using var append = PowerShell.Create(); append.Runspace = runspace;
        append.AddScript("""
            PowerPointSlide -Document $deck -Title Revenue -Layout TitleAndContent { PowerPointParagraph $body -Size 18 }
            PowerPointSlide -Document $deck -Title Margin { PowerPointParagraph 'Margin improved.' }
            """);
        Assert.Empty(append.Invoke()); Assert.False(append.HadErrors, Errors(append));
        using var saved = PowerPointDocument.Load(new MemoryStream(active.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.True(saved.Slides.Count > (cover ? 3 : 2), "The long paragraph must exercise continuation slides.");
        var optional = Assert.Single(saved.GetSections(), section => section.Name == "Optional");
        var results = Assert.Single(saved.GetSections(), section => section.Name == "Results");
        Assert.Equal(boundaries[0].Id, optional.Id); Assert.Empty(optional.SlideIndices);
        Assert.Equal(boundaries[1].Id, results.Id);
        int first = cover ? 1 : 0;
        Assert.Equal(Enumerable.Range(first, saved.Slides.Count - first), results.SlideIndices);
    }

    [Theory]
    [InlineData("remove")]
    [InlineData("duplicate-before-end")]
    [InlineData("move")]
    [InlineData("remove-all")]
    [InlineData("import-before-end")]
    public void PendingSectionSurvivesExistingSlideEditsAndReopeningBeforeItsFirstSlide(string edit) {
        using var document = PowerPointDocument.Create();
        document.AddSlide().AddTitle("Existing first"); document.AddSlide().AddTitle("Existing second");
        using var source = PowerPointDocument.Create(); source.AddSlide().AddTitle("Imported");
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var start = PowerShell.Create(); start.Runspace = runspace;
        start.AddScript("PowerPointSection Results -Document $deck -PassThru");
        var boundary = Assert.IsType<PowerPointSectionInfo>(Assert.Single(start.Invoke()).BaseObject);
        Assert.False(start.HadErrors, Errors(start));
        switch (edit) {
            case "remove": document.RemoveSlide(0); break;
            case "duplicate-before-end": document.DuplicateSlide(0, 1); break;
            case "move": document.MoveSlide(0, 1); break;
            case "remove-all": document.RemoveSlide(1); document.RemoveSlide(0); break;
            case "import-before-end": document.ImportSlide(source, 0, 1); break;
            default: throw new ArgumentException("Unknown edit.", nameof(edit));
        }
        using var reopened = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(reopened.ValidateDocument());
        var pending = Assert.Single(reopened.GetSections(), section => section.Name == "Results");
        Assert.Equal(boundary.Id, pending.Id);
        Assert.Empty(pending.SlideIndices);
        int existing = reopened.Slides.Count;
        runspace.SessionStateProxy.SetVariable("deck", reopened);
        using var append = PowerShell.Create(); append.Runspace = runspace;
        append.AddScript("PowerPointSlide -Document $deck -Title Revenue { PowerPointParagraph 'New results.' }");
        Assert.Empty(append.Invoke());
        Assert.False(append.HadErrors, Errors(append));
        using var saved = PowerPointDocument.Load(new MemoryStream(reopened.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(existing + 1, saved.Slides.Count);
        var results = Assert.Single(saved.GetSections(), section => section.Name == "Results");
        Assert.Equal(boundary.Id, results.Id);
        Assert.Equal(new[] { existing }, results.SlideIndices);
    }
}
