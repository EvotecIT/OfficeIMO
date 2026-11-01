using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlaceholderEditsKeepNativeIdentityGeometryAndOtherIndexesWhenSaved(bool rich) {
        string path = Path.Combine(_root, "edited-placeholder.pptx");
        using (var document = PowerPointDocument.Create()) {
            document.AddSlide().AddTitle("Keep first title");
            var slide = document.AddSlide();
            var body = slide.AddTextBox("Replace this", OfficeLength.Centimeters(2), OfficeLength.Centimeters(4),
                OfficeLength.Centimeters(10), OfficeLength.Centimeters(6));
            body.PlaceholderType = PowerPointPlaceholderType.Body;
            body.PlaceholderIndex = 7;
            var other = slide.AddTextBox("Keep other body");
            other.PlaceholderType = PowerPointPlaceholderType.Body;
            other.PlaceholderIndex = 8;
            document.Save(path);
        }
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        string content = rich ? "{ PowerPointParagraph { PowerPointText 'Status: ' -Bold; PowerPointText 'Healthy' -Italic }; PowerPointParagraph 'All regions' }" : "-Text 'Ready for review'";
        shell.AddScript($$"""
            Open-OfficePowerPoint -Path $target {
                $original = Get-OfficePowerPointPlaceholder -Slide 2 -Type Body -Index 7
                $changed = PowerPointPlaceholder -Slide 2 -Type Body -Index 7 {{content}} -PassThru
                $global:sameHandle = [object]::ReferenceEquals($original, $changed)
                PowerPointPlaceholder -Slide 2 -Type Footer -IgnoreMissing { throw 'Missing placeholder callback must not run' }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.True(Assert.IsType<bool>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("sameHandle")).BaseObject));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal("Keep first title", saved.Slides[0].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
        var selected = saved.Slides[1].GetPlaceholder(PowerPointPlaceholderType.Body, 7)!;
        Assert.Equal(rich ? "Status: Healthy" + Environment.NewLine + "All regions" : "Ready for review", selected.Text);
        Assert.Equal(720000L, selected.Left.ToEmus());
        Assert.Equal(1440000L, selected.Top.ToEmus());
        Assert.Equal(3600000L, selected.Width.ToEmus());
        Assert.Equal(2160000L, selected.Height.ToEmus());
        Assert.Equal("Keep other body", saved.Slides[1].GetPlaceholder(PowerPointPlaceholderType.Body, 8)!.Text);
        if (rich) {
            Assert.Equal(2, selected.Paragraphs.Count);
            Assert.True(selected.Paragraphs[0].Runs[0].Bold);
            Assert.True(selected.Paragraphs[0].Runs[1].Italic);
        }
    }
}
