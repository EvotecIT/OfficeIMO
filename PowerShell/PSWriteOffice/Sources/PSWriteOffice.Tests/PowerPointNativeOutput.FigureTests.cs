using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointNativeOutputWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedTableSideFiguresKeepTheOriginalOnTheFirstPageAndIndependentCopiesLater(bool chart) {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        string imagePath = Path.Combine(_root, "repeated-figure.png");
        File.WriteAllBytes(imagePath, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
        runspace.SessionStateProxy.SetVariable("imagePath", imagePath);
        runspace.SessionStateProxy.SetVariable("asChart", chart);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Table with repeated figure' -Layout TwoContent {
                $rows = 1..8 | ForEach-Object { [pscustomobject]@{ Service = "Service $_"; Count = $_ } }
                PowerPointTable -InputObject $rows -MaxRows 2
                if ($asChart) { $script:figure = PowerPointChart -Title 'Repeated chart' -PassThru }
                else { $script:figure = PowerPointImage -Path $imagePath -PassThru }
                $figure.AltText = 'Repeated context'
                $figure.SetHyperlink([uri]'https://contoso.com/repeated-figure')
                $figure.SetClickSound([IO.MemoryStream]::new([byte[]](1,2,3,4)), 'repeated')
            }
            $figure.AltText = 'First page edited'
            if ($asChart) { $figure.SetTitle('First chart edited') | Out-Null }
            """);
        Assert.Equal(4, deck.Slides.Count);
        var figure = Assert.IsAssignableFrom<PowerPointShape>(System.Management.Automation.PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("figure")).BaseObject);
        if (chart) Assert.Contains(deck.Slides[0].Charts, item => ReferenceEquals(item, figure));
        else Assert.Contains(deck.Slides[0].Pictures, item => ReferenceEquals(item, figure));
        using var saved = PowerPointPresentation.Load(Save(deck, "repeated-native-figure.pptx"));
        for (int i = 0; i < saved.Slides.Count; i++) {
            PowerPointShape stored = chart ? Assert.Single(saved.Slides[i].Charts) : Assert.Single(saved.Slides[i].Pictures);
            Assert.Equal(i == 0 ? "First page edited" : "Repeated context", stored.AltText);
            Assert.Equal("https://contoso.com/repeated-figure", stored.Hyperlink?.AbsoluteUri);
            Assert.Equal(new byte[] { 1, 2, 3, 4 }, stored.GetClickSoundBytes());
            if (chart) {
                Assert.True(((PowerPointChart)stored).TryGetOfficeSnapshot(out var snapshot));
                Assert.Equal(i == 0 ? "First chart edited" : "Repeated chart", snapshot.Title);
            } else Assert.Equal(File.ReadAllBytes(imagePath), ((PowerPointPicture)stored).GetImageBytes());
        }
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CollectedFiguresReturnTheirLiveNativeObjectAndMoveTheirPartsToTheArrangedSlide(bool chart, bool leadingLists) {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        byte[] image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        string imagePath = Path.Combine(_root, "figure.png");
        File.WriteAllBytes(imagePath, image);
        runspace.SessionStateProxy.SetVariable("imagePath", imagePath);
        runspace.SessionStateProxy.SetVariable("asChart", chart);
        runspace.SessionStateProxy.SetVariable("leadingLists", leadingLists);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Native figure' {
                if ($leadingLists) { PowerPointList 'First context'; PowerPointList 'Second context' }
                if ($asChart) {
                    $data = @([pscustomobject]@{ Period='Q1'; Metric=12 }, [pscustomobject]@{ Period='Q2'; Metric=18 }, [pscustomobject]@{ Period='Q3'; Metric=14 })
                    $script:figure = PowerPointChart -InputObject $data -Category Period -Value Metric -Title 'Initial title' -PassThru
                    $figure.SetTitle('Edited before arrangement') | Out-Null
                } else {
                    $script:figure = PowerPointImage -Path $imagePath -PassThru
                }
                $figure.AltText = 'Edited before arrangement'
                $figure.SetHyperlink([uri]'https://contoso.com/before-figure-arrangement')
                $figure.SetClickSound([IO.MemoryStream]::new([byte[]](1,2,3,4)), 'figure')
            }
            $script:arrangedLink = $figure.Hyperlink.AbsoluteUri
            $figure.AltText = 'Edited after arrangement'
            $figure.SetHyperlink([uri]'https://contoso.com/after-figure-arrangement')
            if ($asChart) { $figure.SetTitle('Edited after arrangement') | Out-Null }
            """);
        var figure = Assert.IsAssignableFrom<PowerPointShape>(System.Management.Automation.PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("figure")).BaseObject);
        int index = leadingLists ? 1 : 0;
        Assert.Equal(index + 1, deck.Slides.Count);
        Assert.Equal("https://contoso.com/before-figure-arrangement", Get<string>(runspace, "arrangedLink"));
        if (chart) Assert.Contains(deck.Slides[index].Charts, item => ReferenceEquals(item, figure));
        else Assert.Contains(deck.Slides[index].Pictures, item => ReferenceEquals(item, figure));
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, figure.GetClickSoundBytes());
        using var saved = PowerPointPresentation.Load(Save(deck, "native-figure.pptx"));
        PowerPointShape stored = chart ? Assert.Single(saved.Slides[index].Charts) : Assert.Single(saved.Slides[index].Pictures);
        Assert.Equal("Edited after arrangement", stored.AltText);
        Assert.Equal("https://contoso.com/after-figure-arrangement", stored.Hyperlink?.AbsoluteUri);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, stored.GetClickSoundBytes());
        if (chart) {
            Assert.True(((PowerPointChart)stored).TryGetOfficeSnapshot(out var snapshot));
            Assert.Equal("Edited after arrangement", snapshot.Title);
            Assert.Equal(new[] { 12D, 18D, 14D }, snapshot.Data.Series[0].Values);
        } else Assert.Equal(image, ((PowerPointPicture)stored).GetImageBytes());
        Assert.DoesNotContain(saved.OpenXmlDocument.PresentationPart!.SlideParts.SelectMany(p => p.HyperlinkRelationships),
            link => link.Uri.AbsoluteUri == "https://contoso.com/before-figure-arrangement");
        Assert.Empty(saved.ValidateDocument());
    }
}
