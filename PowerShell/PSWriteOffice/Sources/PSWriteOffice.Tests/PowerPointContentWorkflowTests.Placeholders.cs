using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData("", false)]
    [InlineData("-FromLayout:$false", false)]
    [InlineData("-FromLayout", true)]
    public void PlaceholderReadsSelectNativeTextOrLayoutMetadataAndApplyBothFilters(string mode, bool layout) {
        using var document = PowerPointDocument.Create();
        document.AddSlide(0, 1);
        var selected = document.AddSlide(0, 1);
        var title = selected.AddTitle("Selected second slide", OfficeLength.Points(40), OfficeLength.Points(40),
            OfficeLength.Points(700), OfficeLength.Points(60));
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript($"Get-OfficePowerPointPlaceholder -Document $deck -Slide 2 -Type Title {mode}");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var value = Assert.Single(result).BaseObject;
        if (layout) {
            var info = Assert.IsType<PowerPointLayoutPlaceholderInfo>(value);
            Assert.Equal(PowerPointPlaceholderType.Title, info.PlaceholderType);
            var bounds = Assert.IsType<PowerPointLayoutBox>(info.Bounds);
            Assert.True(bounds.Width.ToEmus() > 0);
        } else {
            Assert.Same(title, value);
            Assert.Equal("Selected second slide", Assert.IsType<PowerPointTextBox>(value).Text);
        }
        shell.Commands.Clear();
        shell.AddScript($"Get-OfficePowerPointPlaceholder -Document $deck -Slide 2 -Type Title -Index 9999 {mode}");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
    }

    [Fact]
    public void PlaceholderPipelineKeepsNativeIdentityAndRejectsAnExplicitForeignDocument() {
        using var document = PowerPointDocument.Create();
        using var other = PowerPointDocument.Create();
        var slide = document.AddSlide(0, 1);
        var title = slide.AddTitle("Native title", OfficeLength.Points(40), OfficeLength.Points(40),
            OfficeLength.Points(700), OfficeLength.Points(60));
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("slide", slide);
        runspace.SessionStateProxy.SetVariable("other", other);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("$slide | Get-OfficePowerPointPlaceholder -Type Title");
        Assert.Same(title, Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        shell.Commands.Clear();
        shell.AddScript("$slide | Get-OfficePowerPointPlaceholder -Document $other -Type Title");
        Assert.Empty(shell.Invoke());
        Assert.True(shell.HadErrors);
        Assert.Contains("does not belong", Errors(shell));
        Assert.Single(document.Slides);
        Assert.Empty(other.Slides);
    }
}
