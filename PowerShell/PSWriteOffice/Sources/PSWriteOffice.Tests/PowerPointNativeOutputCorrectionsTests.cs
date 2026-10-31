using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointNativeOutputWorkflowTests {
    [Theory]
    [InlineData("TitleAndContent")]
    [InlineData("Closing")]
    [InlineData("Auto")]
    public void EmptyListIsValidAndArrangedEvenOnClosing(string layout) {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        runspace.SessionStateProxy.SetVariable("layout", layout);
        Invoke(runspace, """
            1..4 | ForEach-Object { PowerPointSlide -Document $deck -Title "Introduction $_" }
            PowerPointSlide -Document $deck -Title 'Next steps' -Layout $layout {
                $script:list = PowerPointList {} -PassThru
            }
            """);
        var list = Get<PowerPointTextBox>(runspace, "list");
        Assert.Contains(deck.Slides[4].TextBoxes, b => ReferenceEquals(b, list));
        Assert.Equal(PowerPointTextAutoFit.None, list.TextAutoFit);
        using (var saved = PowerPointDocument.Load(Save(deck, "empty-closing.pptx")))
            Assert.Empty(saved.ValidateDocument());
        list.AddParagraph("Added after collection");
        using var edited = PowerPointDocument.Load(Save(deck, "edited-closing.pptx"));
        Assert.Contains(edited.Slides[4].TextBoxes, b => b.Text.Contains("Added after collection", StringComparison.Ordinal));
        Assert.Empty(edited.ValidateDocument());
    }

    [Fact]
    public void CaughtChildFailuresDiscardOnlyTheirNodesAndRelationships() {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Accepted siblings' {
                $script:list = PowerPointList -PassThru {
                    $script:first = PowerPointListItem 'Before' -PassThru
                    try {
                        PowerPointListItem -ErrorAction Stop {
                            $rejected = PowerPointText 'Rejected item' -Link 'https://contoso.com/rejected' -PassThru
                            $rejected.SetClickSound([System.IO.MemoryStream]::new([byte[]](1,2,3,4)), 'discard')
                            throw 'reject-item'
                        }
                    } catch { $script:itemFailed = $true }
                    $script:last = PowerPointListItem -PassThru {
                        PowerPointText 'Accepted ' -Link 'https://contoso.com/accepted'
                        try {
                            PowerPointText 'Rejected run' -Link 'http://[' -ErrorAction Stop
                        } catch { $script:runFailed = $true }
                        try {
                            PowerPointText 'Rejected format' -Size 0 -ErrorAction Stop
                        } catch { $script:formatFailed = $true }
                        PowerPointText 'after failure'
                    }
                }
            }
            """);
        Assert.True(Get<bool>(runspace, "itemFailed"));
        Assert.True(Get<bool>(runspace, "runFailed"));
        Assert.True(Get<bool>(runspace, "formatFailed"));
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "first"));
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "last"));
        var list = Get<PowerPointTextBox>(runspace, "list");
        Assert.Equal(new[] { "Before", "Accepted after failure" }, list.Paragraphs.Select(p => p.Text));
        string path = Save(deck, "child-failures.pptx");
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        using var package = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false);
        var links = package.PresentationPart!.SlideParts.SelectMany(p => p.HyperlinkRelationships).ToList();
        Assert.Equal("https://contoso.com/accepted", Assert.Single(links).Uri.AbsoluteUri);
        Assert.Empty(package.DataParts);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContinuationPreservesBothInteractionSoundsWithoutReauthoring(bool sentences) {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        byte[] click = { 82, 73, 70, 70, 1, 2, 3, 4 };
        byte[] hover = { 82, 73, 70, 70, 5, 6, 7, 8 };
        runspace.SessionStateProxy.SetVariable("click", click);
        runspace.SessionStateProxy.SetVariable("hover", hover);
        runspace.SessionStateProxy.SetVariable("sentences", sentences);
        runspace.SessionStateProxy.SetVariable("narrative", string.Join(" ", Enumerable.Repeat("Repeated sentence with enough detail to require a continuation slide.", 50)) + " ");
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Interaction preservation' {
                if ($sentences) {
                    PowerPointParagraph {
                        PowerPointText $narrative
                        $script:tail = PowerPointText 'Tail interaction.' -PassThru
                        $tail.SetClickSound([System.IO.MemoryStream]::new($click), 'click')
                        $tail.SetMouseOverSound([System.IO.MemoryStream]::new($hover), 'hover')
                    }
                } else {
                    PowerPointList {
                        foreach ($index in 1..42) {
                            PowerPointListItem {
                                $run = PowerPointText "Action $index" -PassThru
                                if ($index -eq 1 -or $index -eq 42) {
                                    $run.SetClickSound([System.IO.MemoryStream]::new($click), 'click')
                                    $run.SetMouseOverSound([System.IO.MemoryStream]::new($hover), 'hover')
                                    if ($index -eq 1) { $script:head = $run }
                                    else { $script:tail = $run; $tail.SetHyperlink('https://contoso.com/tail') }
                                }
                            }
                        }
                    }
                }
            }
            """);
        Assert.True(deck.Slides.Count > 1);
        var tail = Get<PowerPointTextRun>(runspace, "tail");
        Assert.Equal(click, tail.GetClickSoundBytes());
        Assert.Equal(hover, tail.GetMouseOverSoundBytes());
        if (!sentences) {
            var head = Get<PowerPointTextRun>(runspace, "head");
            Assert.Equal(click, head.GetClickSoundBytes());
            Assert.Equal(hover, head.GetMouseOverSoundBytes());
        }
        using var saved = PowerPointDocument.Load(Save(deck, "sounds.pptx"));
        var stored = Assert.Single(saved.Slides.Skip(1).SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs)
            .SelectMany(p => p.Runs), r => r.Text == (sentences ? "Tail interaction." : "Action 42"));
        Assert.Equal(click, stored.GetClickSoundBytes());
        Assert.Equal(hover, stored.GetMouseOverSoundBytes());
        if (!sentences) Assert.Equal("https://contoso.com/tail", stored.Hyperlink!.AbsoluteUri);
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData("Quote", false)]
    [InlineData("Closing", false)]
    [InlineData("Closing", true)]
    public void LargeNativeSpecialLayoutsContinueWithoutClipping(string layout, bool list) {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        string text = string.Join(" ", Enumerable.Repeat("Large authored sentence.", 18));
        runspace.SessionStateProxy.SetVariable("layout", layout);
        runspace.SessionStateProxy.SetVariable("asList", list);
        runspace.SessionStateProxy.SetVariable("longText", text);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Conclusion' -Layout $layout {
                if ($asList) {
                    PowerPointList { $script:paragraph = PowerPointListItem -PassThru { PowerPointText $longText -Size 72 } }
                } else {
                    $script:paragraph = PowerPointParagraph -PassThru { PowerPointText $longText -Size 72 }
                }
            }
            """);
        Assert.True(deck.Slides.Count > 1);
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "paragraph"));
        using var saved = PowerPointDocument.Load(Save(deck, "special-layouts.pptx"));
        Assert.Empty(saved.ValidateDocument());
        Assert.DoesNotContain(saved.Slides.SelectMany(s => s.CreateVisualSnapshot().Diagnostics), d => d.Code == "powerpoint-text-overflow");
        var paragraphs = saved.Slides.SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs).Where(p => p.Runs.Any(r => r.FontSize?.ToPoints() == 72)).ToList();
        Assert.Equal(text, string.Concat(paragraphs.Select(p => p.Text)));
    }

    [Theory]
    [InlineData("Comparison")]
    [InlineData("TwoContent")]
    public void LargeNativeColumnsContinueWithinTheirDesignedFrames(string layout) {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        runspace.SessionStateProxy.SetVariable("layout", layout);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Column fit' -Layout $layout {
                PowerPointColumn -Title 'Before' {
                    $script:list = PowerPointList -Numbered -PassThru {
                        foreach ($i in 1..6) { PowerPointListItem { PowerPointText "Item $i" -Size 72 } }
                    }
                }
                PowerPointColumn -Title 'After' { $script:after = PowerPointParagraph 'Accepted sibling' -PassThru }
            }
            """);
        Assert.True(deck.Slides.Count > 1);
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "after"));
        using var saved = PowerPointDocument.Load(Save(deck, "native-columns.pptx"));
        Assert.Empty(saved.ValidateDocument());
        Assert.DoesNotContain(saved.Slides.SelectMany(s => s.CreateVisualSnapshot().Diagnostics), d => d.Code == "powerpoint-text-overflow");
        var items = saved.Slides.SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs).Where(p => p.Text.StartsWith("Item ", StringComparison.Ordinal)).ToList();
        Assert.Equal(Enumerable.Range(1, 6).Select(i => "Item " + i), items.Select(p => p.Text));
        Assert.All(items.SelectMany(p => p.Runs), r => Assert.Equal(72, r.FontSize?.ToPoints()));
        Assert.Equal(Enumerable.Range(1, 6), items.Select(p => p.NumberingStartAt.GetValueOrDefault()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeLedeSizeIsMeasuredInTheSlideAndColumn(bool column) {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        runspace.SessionStateProxy.SetVariable("column", column);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Lede fit' {
                if ($column) {
                    PowerPointColumn -Title 'Before' {
                        $script:lede = PowerPointParagraph -PassThru { PowerPointText 'Large lede' -Size 72 }
                        PowerPointList 'First', 'Second'
                    }
                    PowerPointColumn -Title 'After' { PowerPointList 'Accepted sibling' }
                } else {
                    $script:lede = PowerPointParagraph -PassThru { PowerPointText 'Large lede' -Size 72 }
                    PowerPointList 'First', 'Second'
                }
            }
            """);
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "lede"));
        using var saved = PowerPointDocument.Load(Save(deck, "native-lede.pptx"));
        Assert.Empty(saved.ValidateDocument());
        Assert.DoesNotContain(saved.Slides.SelectMany(s => s.CreateVisualSnapshot().Diagnostics), d => d.Code == "powerpoint-text-overflow");
    }

    [Fact]
    public void OversizedUnpunctuatedRichParagraphPreservesTextAndUsesMatchingDefaultLeading() {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        string text = string.Join(" ", Enumerable.Repeat("large unpunctuated words", 45));
        runspace.SessionStateProxy.SetVariable("longText", text);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Unpunctuated' -Layout TitleAndContent {
                $script:paragraph = PowerPointParagraph -PassThru {
                    $script:run = PowerPointText $longText -Size 72 -Font Verdana -PassThru
                    PowerPointText ' small tail' -Size 24 -Italic
                }
            }
            """);
        Assert.True(deck.Slides.Count > 1);
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "paragraph"));
        Assert.Same(Node(Get<PowerPointTextRun>(runspace, "run")), Node(Get<PowerPointParagraph>(runspace, "paragraph").Runs[0]));
        using var saved = PowerPointDocument.Load(Save(deck, "long-words.pptx"));
        var bodies = saved.Slides.SelectMany(s => s.TextBoxes).Where(b => b.Paragraphs.Any(p => p.Runs.Any(r => r.FontName == "Verdana"))).ToList();
        Assert.Equal(text + " small tail", string.Concat(bodies.SelectMany(b => b.Paragraphs).Select(p => p.Text)));
        foreach (var paragraph in bodies.SelectMany(b => b.Paragraphs)) {
            double maximum = paragraph.Runs.Max(r => r.FontSize?.ToPoints() ?? 0);
            Assert.True(paragraph.LineSpacing?.ToPoints() >= maximum);
        }
        Assert.All(bodies.SelectMany(b => b.Paragraphs).SelectMany(p => p.Runs).Where(r => r.FontSize?.ToPoints() == 24), r => Assert.True(r.Italic));
        Assert.Empty(saved.ValidateDocument());
        Assert.DoesNotContain(saved.Slides.SelectMany(s => s.CreateVisualSnapshot().Diagnostics), d => d.Code == "powerpoint-text-overflow");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void AuthoredLargeTextAndSpacingDriveContinuationWithoutChangingNativeStyles(bool list, bool nativeEdits) {
        using var deck = PowerPointDocument.Create();
        using var runspace = Open(deck);
        runspace.SessionStateProxy.SetVariable("asList", list);
        runspace.SessionStateProxy.SetVariable("nativeEdits", nativeEdits);
        Invoke(runspace, """
            $size = if ($nativeEdits) { 24 } else { 72 }
            PowerPointSlide -Document $deck -Title 'Large authored text' -Layout TitleAndContent {
                if ($asList) {
                    $script:frame = PowerPointList -PassThru {
                        foreach ($index in 1..9) {
                            $item = PowerPointListItem -PassThru {
                                $run = PowerPointText "Large item $index" -Size $size -Font Verdana -Bold -PassThru
                                if ($nativeEdits) { $run.FontSize = [OfficeIMO.OfficeLength]::Points(72); $run.FontName = 'Verdana' }
                            }
                            $item.SpaceBefore = [OfficeIMO.OfficeLength]::Points(14)
                            $item.SpaceAfter = [OfficeIMO.OfficeLength]::Points(18)
                            $item.LineSpacingMultiplier = 1.3
                        }
                    }
                } else {
                    $script:paragraph = PowerPointParagraph -PassThru {
                        $script:large = PowerPointText ('Large styled sentence. ' * 9).TrimEnd() -Size $size -Font Verdana -Bold -PassThru
                    }
                    if ($nativeEdits) { $large.FontSize = [OfficeIMO.OfficeLength]::Points(72); $large.FontName = 'Verdana' }
                    $paragraph.SpaceBefore = [OfficeIMO.OfficeLength]::Points(14)
                    $paragraph.SpaceAfter = [OfficeIMO.OfficeLength]::Points(18)
                    $paragraph.LineSpacingMultiplier = 1.3
                }
            }
            """);
        Assert.True(deck.Slides.Count > 1, "Authored 72pt content cannot fit a single content frame.");
        if (!list) AssertBacked(deck, Get<PowerPointParagraph>(runspace, "paragraph"));
        using var saved = PowerPointDocument.Load(Save(deck, "large-text.pptx"));
        var frames = saved.Slides.SelectMany(s => s.TextBoxes).Where(b => b.Paragraphs.Any(p => p.Runs.Any(r => r.FontSize?.ToPoints() == 72))).ToList();
        var paragraphs = frames.SelectMany(b => b.Paragraphs).ToList();
        foreach (var paragraph in paragraphs) {
            Assert.All(paragraph.Runs, r => { Assert.Equal(72, r.FontSize?.ToPoints()); Assert.Equal("Verdana", r.FontName); Assert.True(r.Bold); });
            Assert.Equal(14, paragraph.SpaceBefore?.ToPoints());
            Assert.Equal(18, paragraph.SpaceAfter?.ToPoints());
            Assert.Equal(1.3, paragraph.LineSpacingMultiplier);
        }
        foreach (var frame in frames) {
            Assert.Equal(PowerPointTextAutoFit.None, frame.TextAutoFit);
            // Every fixture sentence/item fits in at most two 72pt lines at the native frame width.
            // This lower bound independently catches the original crowded one-frame layout.
            double lowerBound = frame.Paragraphs.Count * (72 * 1.3 + 14 + 18);
            Assert.True(lowerBound <= frame.Height.ToPoints() + 0.1, $"Required at least {lowerBound}pt, frame has {frame.Height.ToPoints()}pt.");
        }
        string content = string.Join(" ", paragraphs.Select(p => p.Text));
        Assert.DoesNotContain(saved.Slides.SelectMany(s => s.CreateVisualSnapshot().Diagnostics), d => d.Code == "powerpoint-text-overflow");
        if (list) foreach (int i in Enumerable.Range(1, 9)) Assert.Contains("Large item " + i, content);
        else Assert.Equal(string.Join(" ", Enumerable.Repeat("Large styled sentence.", 9)), string.Concat(paragraphs.Select(p => p.Text)));
        Assert.Empty(saved.ValidateDocument());
    }
}
