using System.Management.Automation;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData("Modern")]
    [InlineData("Dark")]
    public void NativeFootersRepeatOnceCapturedContentAndKeepSlideOverridesLinksAndFinalCountsAfterAppend(string theme) {
        string path = Path.Combine(_root, "footers.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        runspace.SessionStateProxy.SetVariable("theme", theme);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $script:footerCalls = 0
            New-OfficePowerPoint -Path $target -Theme $theme {
                PowerPointFooter {
                    $script:footerCalls++
                    $script:footerRun = PowerPointText 'Contoso' -Bold -Link 'https://example.com/original' -PassThru
                }
                PowerPointFooter -Alignment Right { PowerPointSlideNumber -IncludeTotal }
                PowerPointSlide -Title 'Review'
                PowerPointSlide -Title 'Decisions' -Layout TitleAndContent {
                    PowerPointFooter 'Local source'
                    foreach ($index in 1..50) {
                        PowerPointParagraph "Decision ${index}: retain the complete evidence and native footer on every continuation."
                    }
                }
                PowerPointSlide -Title 'Summary' -Layout TitleAndContent { PowerPointParagraph 'Ready for discussion.' }
                $script:footerRun.Text = 'Contoso edited'
                $script:footerRun.SetHyperlink('https://example.com/final')
            }
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(1, (int)runspace.SessionStateProxy.GetVariable("footerCalls"));
        int originalCount;
        using (var saved = PowerPointDocument.Load(path)) {
            originalCount = saved.Slides.Count;
            Assert.True(originalCount > 3);
            Assert.Empty(saved.ValidateDocument());
            var firstFooter = Assert.Single(saved.Slides[0].TextBoxes, box => box.Text == "Contoso edited");
            var run = Assert.Single(firstFooter.Paragraphs[0].Runs);
            Assert.True(run.Bold);
            Assert.Equal("https://example.com/final", run.Hyperlink!.AbsoluteUri);
            Assert.All(saved.Slides.Skip(1).Take(originalCount - 2), slide =>
                Assert.Contains(slide.TextBoxes, box => box.Text == "Local source"));
            var lastFooter = Assert.Single(saved.Slides.Last().TextBoxes, box => box.Text == "Contoso");
            Assert.Equal("https://example.com/original", lastFooter.Paragraphs[0].Runs[0].Hyperlink!.AbsoluteUri);
        }
        AssertFooterNumbers(path, originalCount);
        shell.Commands.Clear();
        shell.AddScript("""
            Open-OfficePowerPoint -Path $target {
                PowerPointFooter 'Updated classification'
                PowerPointSlide -Title 'Appendix' -Layout TitleAndContent { PowerPointParagraph 'Appended after reopening.' }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using (var saved = PowerPointDocument.Load(path)) {
            Assert.Equal(originalCount + 1, saved.Slides.Count);
            Assert.Empty(saved.ValidateDocument());
            Assert.Contains(saved.Slides[0].TextBoxes, box => box.Text == "Updated classification");
            var appendedFooter = Assert.Single(saved.Slides.Last().TextBoxes, box => box.Text == "Updated classification");
            Assert.Equal(saved.GetThemeColor(PowerPointThemeColor.Dark2), appendedFooter.Paragraphs[0].Runs[0].Color, ignoreCase: true);
            Assert.All(saved.Slides.Skip(1).Take(originalCount - 2), slide =>
                Assert.Contains(slide.TextBoxes, box => box.Text == "Local source"));
        }
        AssertFooterNumbers(path, originalCount + 1);
    }

    [Fact]
    public void FooterOmissionsKeepTextAndNumbersIndependentAndRejectFieldsOutsideFooterContent() {
        string path = Path.Combine(_root, "footer-omissions.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target {
                PowerPointFooter 'Internal'
                PowerPointFooter -Alignment Right { PowerPointSlideNumber -IncludeTotal }
                PowerPointSlide -Title 'Text only' -Layout TitleAndContent -NoSlideNumber { PowerPointParagraph 'First' }
                PowerPointSlide -Title 'Number only' -Layout TitleAndContent -NoFooter { PowerPointParagraph 'Second' }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using (var saved = PowerPointDocument.Load(path)) {
        Assert.Empty(saved.ValidateDocument());
        Assert.Contains(saved.Slides[0].TextBoxes, box => box.Text == "Internal");
        Assert.DoesNotContain(saved.Slides[0].TextBoxes.SelectMany(box => box.Paragraphs), paragraph => paragraph.Text.Contains(" / "));
        Assert.DoesNotContain(saved.Slides[1].TextBoxes, box => box.Text == "Internal");
        Assert.Contains(saved.Slides[1].TextBoxes, box => box.Text == "2");
        }
        shell.Commands.Clear();
        shell.AddScript("Open-OfficePowerPoint -Path $target { PowerPointSlide -Title 'Appended' -Layout TitleAndContent { PowerPointParagraph 'Third' } }");
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using (var reopened = PowerPointDocument.Load(path)) {
            Assert.Empty(reopened.ValidateDocument());
            Assert.Contains(reopened.Slides[2].TextBoxes, box => box.Text == "3 / 3");
            Assert.Contains(reopened.Slides[2].TextBoxes, box => box.Text == "Internal");
            Assert.DoesNotContain(reopened.Slides[0].TextBoxes.SelectMany(box => box.Paragraphs), paragraph => paragraph.Text.Contains(" / "));
        }
        shell.Commands.Clear();
        shell.AddScript("PowerPointSlideNumber -IncludeTotal -ErrorAction Stop");
        var rejected = shell.Invoke();
        Assert.Empty(rejected);
        Assert.True(shell.HadErrors);
        Assert.Contains(shell.Streams.Error, error => error.ToString().Contains("inside PowerPointFooter"));
    }

    [Fact]
    public void ASlideOnlyFooterRemainsLocalWhenADeckFooterIsAddedAfterReopening() {
        string path = Path.Combine(_root, "local-footer.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target {
                PowerPointSlide -Title 'Local evidence' -Layout TitleAndContent {
                    PowerPointFooter 'Local only'
                    PowerPointParagraph 'First'
                }
            }
            Open-OfficePowerPoint -Path $target {
                PowerPointFooter 'Deck default'
                PowerPointSlide -Title 'New evidence' -Layout TitleAndContent { PowerPointParagraph 'Second' }
            }
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Contains(saved.Slides[0].TextBoxes, box => box.Text == "Local only");
        Assert.DoesNotContain(saved.Slides[0].TextBoxes, box => box.Text == "Deck default");
        Assert.Contains(saved.Slides[1].TextBoxes, box => box.Text == "Deck default");
    }

    private static void AssertFooterNumbers(string path, int total) {
        using var package = PresentationDocument.Open(path, false);
        var part = package.PresentationPart!;
        var slides = part.Presentation!.SlideIdList!.Elements<P.SlideId>()
            .Select(id => ((SlidePart)part.GetPartById(id.RelationshipId!)).Slide!).ToArray();
        Assert.Equal(total, slides.Length);
        for (int index = 0; index < slides.Length; index++) {
            var visible = slides[index].CommonSlideData!.ShapeTree!;
            var field = Assert.Single(visible.Descendants<A.Field>(), item => item.Type?.Value == "slidenum");
            Assert.Equal((index + 1).ToString(), field.Text!.Text);
            Assert.Contains(visible.Descendants<A.Run>(), run => run.Text?.Text == " / " + total);
            Assert.Contains(visible.Descendants<P.PlaceholderShape>(), placeholder => placeholder.Type?.Value == P.PlaceholderValues.SlideNumber);
        }
    }
}
