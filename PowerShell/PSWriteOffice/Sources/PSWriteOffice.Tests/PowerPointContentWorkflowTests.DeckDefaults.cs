using System.Globalization;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeckDefaultsSurviveNativeSaveAndPhysicalContinuations(bool suppressChrome) {
        string path = Path.Combine(_root, "defaults.pptx"), copied = Path.Combine(_root, "copied.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        runspace.SessionStateProxy.SetVariable("copied", copied);
        runspace.SessionStateProxy.SetVariable("suppress", suppressChrome);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target -Theme Dark -Author 'Review team' -Date ([datetime]'2026-10-30') `
                -Accent '#008040' -Seed 'review-2026' -SlideSize Screen4x3 -NoFooter:$suppress -NoSlideNumber:$suppress `
                -ContinuationFormat 'Part {1} of {2}: {0}' {
                PowerPointSlide -Layout Title -Title 'Service review'
                PowerPointSlide -Layout TitleAndContent -Title 'Decisions' {
                    PowerPointParagraph (('A decision needs an owner and a date. ' * 120).Trim())
                }
            }
            New-OfficePowerPoint -Path $copied -Template $target -Accent '#A04000' {
                PowerPointSlide -Layout TitleAndContent -Title 'Added to template' { PowerPointParagraph 'Keep the template size.' }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal("Review team", saved.BuiltinDocumentProperties.Creator);
        Assert.Equal(OfficeLength.Inches(10), saved.SlideSize.Width);
        Assert.Equal(OfficeLength.Inches(7.5), saved.SlideSize.Height);
        Assert.Equal("008040", saved.GetTheme().Palette.Accent1.ToRgbHex());
        Assert.Equal(OfficeTheme.Dark.Palette.Background, saved.GetTheme().Palette.Background);
        Assert.Equal(OfficeTheme.Dark.Fonts.Body.Family, saved.GetTheme().Fonts.Body.Family);
        string cover = string.Join("\n", saved.Slides[0].TextBoxes.Select(box => box.Text));
        Assert.Contains("Review team", cover);
        Assert.Contains(new DateTime(2026, 10, 30).ToString("d", CultureInfo.CurrentCulture), cover);
        Assert.True(saved.Slides.Count > 2, "The paragraph must exercise physical continuations.");
        Assert.Contains(saved.Slides.Skip(2).SelectMany(slide => slide.TextBoxes), box => box.Text == $"Part 2 of {saved.Slides.Count - 1}: Decisions");
        Assert.All(saved.Slides.Skip(1), slide => {
            Assert.Equal(!suppressChrome, slide.TextBoxes.Any(box => box.ShapePlaceholderType == PowerPointPlaceholderType.Footer));
            Assert.Equal(!suppressChrome, slide.TextBoxes.Any(box => box.ShapePlaceholderType == PowerPointPlaceholderType.SlideNumber));
        });
        using var templateCopy = PowerPointDocument.Load(copied);
        Assert.Empty(templateCopy.ValidateDocument());
        Assert.Equal(saved.SlideSize.Width, templateCopy.SlideSize.Width);
        Assert.Equal(saved.SlideSize.Height, templateCopy.SlideSize.Height);
        Assert.Equal(saved.Slides.Count + 1, templateCopy.Slides.Count);
        Assert.Equal("A04000", templateCopy.GetTheme().Palette.Accent1.ToRgbHex());
        Assert.Equal(saved.GetTheme().Palette.Background, templateCopy.GetTheme().Palette.Background);
        Assert.Equal(saved.GetTheme().Fonts.Body.Family, templateCopy.GetTheme().Fonts.Body.Family);
    }

    [Theory]
    [InlineData("-SlideSize 9999")]
    [InlineData("-ContinuationFormat '{3}'")]
    public void InvalidDeckDefaultsPreserveTheDestinationAndDoNotRunContent(string arguments) {
        string path = Path.Combine(_root, "existing.pptx");
        byte[] original = { 1, 2, 3, 4 };
        File.WriteAllBytes(path, original);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("$global:entered = $false; New-OfficePowerPoint -Path $target -Force " + arguments + " { $global:entered = $true }; $global:entered");
        Assert.False(Assert.IsType<bool>(Assert.Single(shell.Invoke()).BaseObject));
        Assert.True(shell.HadErrors);
        Assert.Equal(original, File.ReadAllBytes(path));
    }
}
