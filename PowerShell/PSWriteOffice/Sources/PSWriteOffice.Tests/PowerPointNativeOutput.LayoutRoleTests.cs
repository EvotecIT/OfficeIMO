using OfficeIMO;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointNativeOutputWorkflowTests {
    [Theory]
    [InlineData("Closing", false, false)]
    [InlineData("Closing", true, false)]
    [InlineData("Auto", false, false)]
    [InlineData("Auto", true, false)]
    [InlineData("Closing", false, true)]
    [InlineData("Auto", false, true)]
    public void ClosingKeepsBothContactAndSiblingList(string layout, bool empty, bool longContact) {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        string contact = longContact ? string.Concat(Enumerable.Repeat("Contact retained. ", 25)) : "Contact retained";
        runspace.SessionStateProxy.SetVariable("layout", layout);
        runspace.SessionStateProxy.SetVariable("empty", empty);
        runspace.SessionStateProxy.SetVariable("contact", contact);
        runspace.SessionStateProxy.SetVariable("size", longContact ? 72 : 16);
        Invoke(runspace, """
            1..4 | ForEach-Object { PowerPointSlide -Document $deck -Title "Introduction $_" }
            PowerPointSlide -Document $deck -Title 'Conclusion' -Layout $layout {
                $script:contactParagraph = PowerPointParagraph -PassThru { PowerPointText $contact -Size $size }
                $script:siblingList = PowerPointList -PassThru {
                    if (-not $empty) { PowerPointListItem 'Accepted step' }
                }
            }
            """);
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "contactParagraph"));
        var list = Get<PowerPointTextBox>(runspace, "siblingList");
        Assert.Contains(deck.Slides.SelectMany(slide => slide.TextBoxes), box => ReferenceEquals(box, list));
        Assert.True(list.Left.ToPoints() > deck.SlideSize.Width.ToEmus() / 12700D / 2D, "The accepted list must occupy the Closing step frame.");
        Assert.Equal(PowerPointTextAutoFit.None, list.TextAutoFit);
        using var saved = PowerPointDocument.Load(Save(deck, "closing-siblings.pptx"));
        Assert.Empty(saved.ValidateDocument());
        Assert.DoesNotContain(saved.Slides.SelectMany(slide => slide.CreateVisualSnapshot().Diagnostics), d => d.Code == "powerpoint-text-overflow");
        var paragraphs = saved.Slides.SelectMany(slide => slide.TextBoxes).SelectMany(box => box.Paragraphs).ToList();
        Assert.Equal(contact, string.Concat(paragraphs.Where(p => p.Runs.Any(r => r.FontSize?.ToPoints() == (longContact ? 72 : 16))).Select(p => p.Text)));
        Assert.Equal(empty ? 0 : 1, paragraphs.Count(p => p.Text == "Accepted step"));
    }

    [Fact]
    public void QuoteContinuationMeasuresTheInheritedHeadingFont() {
        var fonts = OfficeTheme.Modern.Fonts.With(heading: new OfficeThemeFont("Verdana", Array.Empty<string>()),
            body: new OfficeThemeFont("Calibri", Array.Empty<string>()));
        using var deck = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern.WithFonts(fonts) });
        using var runspace = Open(deck);
        string text = string.Concat(Enumerable.Repeat("Wide WWW quote words ", 25));
        runspace.SessionStateProxy.SetVariable("quoteText", text);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Quotation' -Layout Quote {
                $script:quoteParagraph = PowerPointParagraph -PassThru { PowerPointText $quoteText -Size 72 }
                $quoteParagraph.LineSpacing = [OfficeIMO.OfficeLength]::Points(90)
            }
            """);
        Assert.True(deck.Slides.Count > 1);
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "quoteParagraph"));
        using var saved = PowerPointDocument.Load(Save(deck, "quote-role-font.pptx"));
        Assert.Empty(saved.ValidateDocument());
        var paragraphs = saved.Slides.SelectMany(slide => slide.TextBoxes).SelectMany(box => box.Paragraphs)
            .Where(p => p.Runs.Any(run => run.FontSize?.ToPoints() == 72)).ToList();
        Assert.Equal(text, string.Concat(paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal("Verdana", saved.GetThemeFonts().MajorLatin);
        Assert.All(paragraphs.SelectMany(paragraph => paragraph.Runs), run => Assert.Contains(run.FontName, new[] { "Verdana", "+mj-lt" }));
        Assert.DoesNotContain(saved.Slides.SelectMany(slide => slide.CreateVisualSnapshot().Diagnostics), d => d.Code == "powerpoint-text-overflow");
    }
}
