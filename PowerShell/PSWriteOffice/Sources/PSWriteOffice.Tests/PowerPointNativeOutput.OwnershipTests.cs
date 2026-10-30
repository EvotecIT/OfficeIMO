using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointNativeOutputWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RejectedFirstNotesRemoveTheirAudioFromTheLivePackage(bool wholeSlide) {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        runspace.SessionStateProxy.SetVariable("wholeSlide", wholeSlide);
        Invoke(runspace, """
            $rejectedContent = {
                $notes = PowerPointNotes 'Rejected first notes' -PassThru
                $run = $notes.Notes.Paragraphs[0].Runs[0]
                $run.SetClickSound([System.IO.MemoryStream]::new([byte[]](1,2,3,4)), 'rejected-click')
                $run.SetMouseOverSound([System.IO.MemoryStream]::new([byte[]](5,6,7,8)), 'rejected-hover')
                throw 'reject notes'
            }
            $script:leaked = $false
            if ($wholeSlide) {
                foreach ($attempt in 1..3) {
                    try { PowerPointSlide -Document $deck -Title 'Rejected' -ErrorAction Stop { & $rejectedContent } } catch { }
                    $script:leaked = $script:leaked -or @($deck.OpenXmlDocument.DataParts).Count -ne 0
                }
                PowerPointSlide -Document $deck -Title 'Accepted' { PowerPointParagraph 'Kept' }
            } else {
                PowerPointSlide -Document $deck -Title 'Accepted' {
                    foreach ($attempt in 1..3) {
                        try { PowerPointColumn -ErrorAction Stop { & $rejectedContent } } catch { }
                        $script:leaked = $script:leaked -or @($deck.OpenXmlDocument.DataParts).Count -ne 0
                    }
                    PowerPointParagraph 'Kept'
                }
            }
            """);
        Assert.False(Get<bool>(runspace, "leaked"), "Rejected notes must release their audio before another operation starts.");
        Assert.Empty(deck.OpenXmlDocument.DataParts);
        Assert.Single(deck.Slides);
        Assert.Null(deck.OpenXmlDocument.PresentationPart!.SlideParts.Single().NotesSlidePart);
        using var saved = PowerPointPresentation.Load(Save(deck, "rejected-first-notes.pptx"));
        Assert.Empty(saved.OpenXmlDocument.DataParts);
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void RejectedColumnRestoresNotesLinksAndAudioAndRemovesRejectedRelationships() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        byte[] original = { 82, 73, 70, 70, 1, 2, 3, 4 };
        runspace.SessionStateProxy.SetVariable("audio", original);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Notes transaction' {
                $notes = PowerPointNotes 'Original' -PassThru
                $run = $notes.Notes.Paragraphs[0].Runs[0]
                $run.SetHyperlink('https://contoso.com/original')
                $run.SetClickSound([System.IO.MemoryStream]::new($audio), 'original')
                try {
                    PowerPointColumn -ErrorAction Stop {
                        $run.ClearHyperlink()
                        $run.SetClickSound([System.IO.MemoryStream]::new([byte[]](5,6,7,8)), 'rejected')
                        $run.SetHyperlink('https://contoso.com/rejected')
                        $run.SetMouseOverSound([System.IO.MemoryStream]::new([byte[]](9,10,11,12)), 'rejected-hover')
                        throw 'reject column'
                    }
                } catch { $script:rejected = $true }
                PowerPointParagraph 'Accepted'
            }
            """);
        Assert.True(Get<bool>(runspace, "rejected"));
        string path = Save(deck, "notes-relationships.pptx");
        using var saved = PowerPointPresentation.Load(path);
        var stored = Assert.Single(saved.Slides[0].Notes.Paragraphs).Runs[0];
        Assert.Equal("https://contoso.com/original", stored.Hyperlink?.AbsoluteUri);
        Assert.Equal(original, stored.GetClickSoundBytes());
        Assert.Null(stored.GetMouseOverSoundBytes());
        using var package = PresentationDocument.Open(path, false);
        var notes = package.PresentationPart!.SlideParts.Single().NotesSlidePart!;
        Assert.Equal("https://contoso.com/original", Assert.Single(notes.HyperlinkRelationships).Uri.AbsoluteUri);
        Assert.Single(notes.DataPartReferenceRelationships);
        Assert.Single(package.DataParts);
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void ArrangingNativeFramePreservesAuthoredAppearanceAndBodyProperties() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Authored frame' {
                $script:list = PowerPointList 'Ready' -PassThru
                $list.FillColor = '127733'
                $list.SetShadow('112233')
                $list.TextDirection = [OfficeIMO.PowerPoint.PowerPointTextDirection]::Vertical
            }
            """);
        var list = Get<PowerPointTextBox>(runspace, "list");
        Assert.Contains(deck.Slides[0].TextBoxes, b => ReferenceEquals(b, list));
        using var saved = PowerPointPresentation.Load(Save(deck, "authored-frame.pptx"));
        var stored = Assert.Single(saved.Slides[0].TextBoxes, b => b.Text.Contains("Ready", StringComparison.Ordinal));
        Assert.Equal("127733", stored.FillColor);
        Assert.Equal(PowerPointTextDirection.Vertical, stored.TextDirection);
        using var package = PresentationDocument.Open(Path.Combine(_root, "authored-frame.pptx"), false);
        var shape = Assert.Single(package.PresentationPart!.SlideParts.Single().Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Shape>(),
            s => s.TextBody?.InnerText == "Ready");
        Assert.Single(shape.ShapeProperties!.Descendants<DocumentFormat.OpenXml.Drawing.OuterShadow>());
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void FrameInteractionRelationshipsFollowTheFrameToItsContinuationSlide() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        byte[] audio = { 82, 73, 70, 70, 1, 2, 3, 4 };
        runspace.SessionStateProxy.SetVariable("audio", audio);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Three action lists' {
                PowerPointList ('First ' * 12), ('Second ' * 12)
                PowerPointList ('Third ' * 12), ('Fourth ' * 12)
                $script:last = PowerPointList 'Last live frame' -PassThru
                $last.SetHyperlink([uri]'https://contoso.com/frame')
                $last.SetClickSound([System.IO.MemoryStream]::new($audio), 'frame')
            }
            """);
        var frame = Get<PowerPointTextBox>(runspace, "last");
        Assert.Contains(deck.Slides.Skip(1).SelectMany(s => s.TextBoxes), b => ReferenceEquals(b, frame));
        Assert.Equal("https://contoso.com/frame", frame.Hyperlink?.AbsoluteUri);
        Assert.Equal(audio, frame.GetClickSoundBytes());
        using var saved = PowerPointPresentation.Load(Save(deck, "frame-interaction.pptx"));
        var stored = Assert.Single(saved.Slides.Skip(1).SelectMany(s => s.TextBoxes), b => b.Text == "Last live frame");
        Assert.Equal("https://contoso.com/frame", stored.Hyperlink?.AbsoluteUri);
        Assert.Equal(audio, stored.GetClickSoundBytes());
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void AuthoredWideQuoteUsesTheWidthMeasuredBeforeContinuation() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        string text = string.Join(" ", Enumerable.Repeat("Wide words fill this authored quote", 30));
        runspace.SessionStateProxy.SetVariable("quoteText", text);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Authored quote' -Layout Quote {
                PowerPointParagraph { $script:quoteRun = PowerPointText $quoteText -Size 48 -Font 'Courier New' -PassThru }
            }
            """);
        using var saved = PowerPointPresentation.Load(Save(deck, "wide-quote.pptx"));
        Assert.True(saved.Slides.Count > 1);
        Assert.Equal(text, string.Concat(saved.Slides.SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs)
            .Where(p => p.Runs.Any(r => r.FontName == "Courier New")).Select(p => p.Text)));
        Assert.DoesNotContain(saved.Slides.SelectMany(s => s.CreateVisualSnapshot().Diagnostics), d => d.Code == "powerpoint-text-overflow");
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void RepeatedTableSideListKeepsItsFirstNativeNodesAndClonesLaterPages() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Table with context' -Layout TwoContent {
                $rows = 1..8 | ForEach-Object { [pscustomobject]@{ Service = "Service $_"; Count = $_ } }
                PowerPointTable -InputObject $rows -MaxRows 2
                $script:context = PowerPointList 'Retained context' -PassThru
                $script:contextRun = $context.Paragraphs[0].Runs[0]
                $contextRun.SetHyperlink('https://contoso.com/context')
            }
            $contextRun.Text = 'First page context edited'
            """);
        Assert.Equal(4, deck.Slides.Count);
        var frame = Get<PowerPointTextBox>(runspace, "context");
        Assert.Contains(deck.Slides[0].TextBoxes, b => ReferenceEquals(b, frame));
        using var saved = PowerPointPresentation.Load(Save(deck, "table-side-list.pptx"));
        Assert.Contains(saved.Slides[0].TextBoxes, b => b.Text == "First page context edited");
        Assert.All(saved.Slides.Skip(1), s => Assert.Contains(s.TextBoxes, b => b.Text == "Retained context"));
        Assert.All(saved.Slides, s => Assert.Contains(s.TextBoxes.SelectMany(b => b.Paragraphs).SelectMany(p => p.Runs),
            r => r.Hyperlink?.AbsoluteUri == "https://contoso.com/context"));
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CollectedTableReturnsItsNativeTableAndPreservesLiveCellEditsAndStyleAcrossContinuation(bool leadingLists) {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        const string style = "{5940675A-B579-460E-94D1-54222C63F5DA}";
        runspace.SessionStateProxy.SetVariable("tableStyle", style);
        runspace.SessionStateProxy.SetVariable("leadingLists", leadingLists);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Editable table' {
                if ($leadingLists) { PowerPointList 'First context'; PowerPointList 'Second context' }
                $rows = 1..8 | ForEach-Object { [pscustomobject]@{ Service = "Service $_"; Count = $_ } }
                $script:table = PowerPointTable -InputObject $rows -MaxRows 2 -StyleId $tableStyle -PassThru
                $table.GetCell(1,0).Text = 'Edited before arrangement'
                $script:lastCell = $table.GetCell(8,0)
                $lastCell.Paragraphs[0].Runs[0].SetHyperlink('https://contoso.com/before-table-continuation')
            }
            $script:beforeLink = $lastCell.Paragraphs[0].Runs[0].Hyperlink.AbsoluteUri
            $table.GetCell(1,0).Text = 'Edited after arrangement'
            $lastCell.Text = 'Edited on last continuation'
            $lastCell.Paragraphs[0].Runs[0].SetHyperlink('https://contoso.com/after-table-continuation')
            """);
        var table = Get<PowerPointTable>(runspace, "table");
        Assert.Contains(deck.Slides[leadingLists ? 1 : 0].Tables, t => ReferenceEquals(t, table));
        Assert.Equal(leadingLists ? 5 : 4, deck.Slides.Count);
        Assert.Equal("https://contoso.com/before-table-continuation", Get<string>(runspace, "beforeLink"));
        using var saved = PowerPointPresentation.Load(Save(deck, "native-table.pptx"));
        Assert.Equal("Edited after arrangement", saved.Slides[leadingLists ? 1 : 0].Tables.Single().GetCell(1, 0).Text);
        Assert.Equal("Edited on last continuation", saved.Slides.Last().Tables.Single().GetCell(2, 0).Text);
        Assert.Equal("https://contoso.com/after-table-continuation", saved.Slides.Last().Tables.Single().GetCell(2, 0).Paragraphs[0].Runs[0].Hyperlink?.AbsoluteUri);
        Assert.All(saved.Slides.Where(s => s.Tables.Any()), s => Assert.Equal(style, s.Tables.Single().StyleId));
        Assert.Empty(saved.ValidateDocument());
    }
}
