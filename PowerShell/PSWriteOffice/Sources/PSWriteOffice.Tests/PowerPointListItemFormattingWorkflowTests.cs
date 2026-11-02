using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointNativeOutputWorkflowTests {
    [Fact]
    public void PlainListItemFormattingUsesNativeRunsAndSurvivesReopening() {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Escalation' {
                PowerPointList {
                    $script:item = PowerPointListItem 'Escalate capacity' -Level 2 -Bold -Italic -Strikethrough -Underline Double -Color Accent2 -Highlight '#FEF3C7' -Font Aptos -Size '0.25in' -PassThru
                }
            }
            """);
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "item"));
        using var saved = PowerPointDocument.Load(Save(deck, "list-item-format.pptx"));
        Assert.Empty(saved.ValidateDocument());
        var paragraph = Assert.Single(saved.Slides.SelectMany(slide => slide.TextBoxes).SelectMany(box => box.Paragraphs),
            item => item.Text == "Escalate capacity");
        Assert.Equal(1, paragraph.Level);
        var run = Assert.Single(paragraph.Runs);
        Assert.True(run.Bold); Assert.True(run.Italic); Assert.True(run.Strikethrough);
        Assert.Equal(PowerPointUnderlineStyle.Double, run.UnderlineStyle);
        Assert.Equal(saved.GetThemeColor(PowerPointThemeColor.Accent2), run.Color);
        Assert.Equal("FEF3C7", run.HighlightColor);
        Assert.Equal("Aptos", run.FontName);
        Assert.Equal(18D, run.FontSize?.ToPoints());
    }

    [Fact]
    public void RejectedPlainItemStyleDoesNotRetainAPartiallyFormattedParagraph() {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Accepted items' {
                PowerPointList {
                    PowerPointListItem 'Before'
                    try { PowerPointListItem 'Rejected style' -Bold -Size 0 -ErrorAction Stop } catch { }
                    PowerPointListItem 'After' -Italic
                }
            }
            """);
        using var saved = PowerPointDocument.Load(Save(deck, "list-item-format-rejected.pptx"));
        Assert.Empty(saved.ValidateDocument());
        var paragraphs = saved.Slides.SelectMany(slide => slide.TextBoxes).SelectMany(box => box.Paragraphs).ToList();
        Assert.Contains(paragraphs, item => item.Text == "Before");
        Assert.Contains(paragraphs, item => item.Text == "After");
        Assert.DoesNotContain(paragraphs, item => item.Text == "Rejected style");
    }
}
