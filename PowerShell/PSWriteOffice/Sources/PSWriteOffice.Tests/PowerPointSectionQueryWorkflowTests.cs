using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void SectionIndicesConvertToOneBasedNumbersForRemovingTheSelectedSectionOnly() {
        using var document = PowerPointDocument.Create();
        document.AddSlide().AddTitle("Keep first");
        document.AddSlide().AddTitle("Keep second");
        document.AddSlide().AddTitle("Appendix first");
        document.AddSlide().AddTitle("Appendix second");
        document.AddSection("Results", 0);
        document.AddSection("Appendix", 2);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            $appendix = Get-OfficePowerPointSection -Document $deck -Name Appendix
            $slides = @($appendix.SlideIndices | ForEach-Object { $_ + 1 })
            Remove-OfficePowerPointSlide -Document $deck -Slide $slides -Confirm:$false
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var reopened = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(reopened.ValidateDocument());
        Assert.Equal(new[] { "Keep first", "Keep second" }, reopened.Slides.Select(slide =>
            slide.GetPlaceholder(PowerPointPlaceholderType.Title)!.Text));
        Assert.Equal("Results", Assert.Single(reopened.GetSections()).Name);
    }

    [Theory]
    [InlineData("-Name 'app*'", "Appendix A|Appendix B")]
    [InlineData("-Name 'app*' -CaseSensitive", "")]
    [InlineData("-Name 'Appendix ?' -CaseSensitive", "Appendix A|Appendix B")]
    [InlineData("-Name 'Notes`[1`]$'", "Notes[1]$")]
    public void SectionNamePatternsUsePowerShellWildcardsAndExplicitCaseSensitivity(string arguments, string expected) {
        using var document = PowerPointDocument.Create();
        document.AddSlide(); document.AddSlide(); document.AddSlide();
        document.AddSection("Appendix A", 0);
        document.AddSection("Appendix B", 1);
        document.AddSection("Notes[1]$", 2);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("$deck | Get-OfficePowerPointSection " + arguments);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(expected, string.Join("|", result.Select(item => Assert.IsType<PowerPointSectionInfo>(item.BaseObject).Name)));
        Assert.Equal(new[] { "Appendix A", "Appendix B", "Notes[1]$" }, document.GetSections().Select(section => section.Name));
    }
}
