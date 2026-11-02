using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData("-Slide $slide", false)]
    [InlineData("", true)]
    public void PlacedParagraphsAndListsKeepNativeEditsFormattingAndPhysicalBounds(string selection, bool pipeline) {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("slide", slide);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        string prefix = pipeline ? "$slide | " : string.Empty;
        shell.AddScript("$paragraph = " + prefix + "PowerPointParagraph " + selection + """
             -X 2cm -Y 3cm -Width 12cm -Height 1cm -PassThru {
                PowerPointText 'Status' -Bold
                PowerPointText ' ready' -Italic
            }
            $paragraph.AddText(' updated') | Out-Null
            """ + "\n$list = " + prefix + "PowerPointList " + selection + """
             -Column 7 -Span 6 -Y 5cm -Height 2cm -Numbering RomanUpperCharacterPeriod -PassThru {
                PowerPointListItem 'First action'
                PowerPointListItem { PowerPointText 'Second action' -Bold }
            }
            $list.Paragraphs[1].AddText(' updated') | Out-Null
            $paragraph
            $list
            """);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(2, output.Count);
        var paragraph = Assert.IsType<PowerPointParagraph>(output[0].BaseObject);
        var list = Assert.IsType<PowerPointTextBox>(output[1].BaseObject);
        Assert.Equal("Status ready updated", paragraph.Text);
        Assert.Contains(slide.TextBoxes, box => ReferenceEquals(box, list));
        string path = Path.Combine(_root, "placed-paragraph-list.pptx");
        document.Save(path);
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var savedSlide = Assert.Single(saved.Slides);
        var paragraphBox = Assert.Single(savedSlide.TextBoxes, box => box.Text == "Status ready updated");
        Assert.Equal(720000L, paragraphBox.Left.ToEmus());
        Assert.Equal(1080000L, paragraphBox.Top.ToEmus());
        Assert.Equal(4320000L, paragraphBox.Width.ToEmus());
        Assert.Equal(360000L, paragraphBox.Height.ToEmus());
        var savedParagraph = Assert.Single(paragraphBox.Paragraphs);
        Assert.True(savedParagraph.Runs[0].Bold);
        Assert.True(savedParagraph.Runs[1].Italic);
        var savedList = Assert.Single(savedSlide.TextBoxes, box => box.Paragraphs.Count == 2);
        Assert.Equal(new[] { "First action", "Second action updated" }, savedList.Paragraphs.Select(item => item.Text));
        Assert.All(savedList.Paragraphs, item => Assert.Equal(PowerPointNumberingScheme.RomanUpperCharacterPeriod, item.NumberingScheme));
        Assert.Equal(new int?[] { 1, 2 }, savedList.Paragraphs.Select(item => item.NumberingStartAt));
        Assert.Equal(1800000L, savedList.Top.ToEmus());
        Assert.Equal(720000L, savedList.Height.ToEmus());
        Assert.True(savedList.Left > paragraphBox.Left);
        Assert.True(savedList.Width.ToEmus() > 0);
        Assert.True(savedList.Paragraphs[1].Runs[0].Bold);
    }
}
