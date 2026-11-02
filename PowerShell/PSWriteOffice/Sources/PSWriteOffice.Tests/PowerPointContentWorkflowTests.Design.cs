using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void DeckDesignOptionsChangeSavedDensityAndGeometryWithoutReplacingTheSelectedTheme() {
        string image = Path.Combine(_root, "evidence.svg");
        File.WriteAllText(image, ThreeBandSvg);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("image", image);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        const string note = "Traffic increased while the service remained healthy.";
        var results = new Dictionary<string, (double Size, double TextWidth, double ImageWidth, double Left, double Right)>();
        foreach (var item in new[] {
                     ("Default", ""), ("ExplicitDefault", "-LayoutStrategy ContentFirst"), ("Compact", "-LayoutStrategy Compact"),
                     ("Visual", "-LayoutStrategy VisualFirst"), ("Style", "-Style QuietAppendix"),
                     ("Purpose", "-Purpose 'technical proposal'"),
                     ("Design", "-LayoutStrategy DesignFirst"), ("Repeat", "-LayoutStrategy DesignFirst") }) {
            string path = Path.Combine(_root, item.Item1 + ".pptx");
            runspace.SessionStateProxy.SetVariable("target", path);
            runspace.SessionStateProxy.SetVariable("seed", item.Item1 is "Design" or "Repeat" ? "campaign9" : "service-review");
            shell.Commands.Clear();
            shell.AddScript($$"""
                New-OfficePowerPoint -Path $target -Theme Dark -Seed $seed {{item.Item2}} {
                    PowerPointSlide -Title 'Service review'
                    PowerPointSlide -Title 'Capacity and demand' -Layout TwoContent {
                        PowerPointColumn { PowerPointParagraph '{{note}}' }
                        PowerPointImage -Path $image
                    }
                    PowerPointSlide -Title 'Decisions' -Layout TwoContent {
                        PowerPointColumn { PowerPointParagraph 'Keep the complete review evidence.' }
                        PowerPointColumn { PowerPointParagraph 'Track the next action to closure.' }
                    }
                }
                """);
            Assert.Empty(shell.Invoke());
            Assert.False(shell.HadErrors, Errors(shell));
            using var saved = PowerPointDocument.Load(path);
            Assert.Empty(saved.ValidateDocument());
            Assert.Equal(3, saved.Slides.Count);
            Assert.Equal(OfficeTheme.Dark.Palette.Background.ToRgbHex(), saved.GetThemeColor(PowerPointThemeColor.Light1));
            Assert.Equal(OfficeTheme.Dark.Palette.Text.ToRgbHex(), saved.GetThemeColor(PowerPointThemeColor.Dark1));
            Assert.Equal(OfficeTheme.Dark.Palette.Accent1.ToRgbHex(), saved.GetThemeColor(PowerPointThemeColor.Accent1));
            Assert.Equal(OfficeTheme.Dark.Fonts.Body.Family, saved.GetThemeFonts().MinorLatin);
            var body = Assert.Single(saved.Slides[1].TextBoxes, box => box.Text == note);
            var picture = Assert.Single(saved.Slides[1].Pictures);
            var left = Assert.Single(saved.Slides[2].TextBoxes, box => box.Text == "Keep the complete review evidence.");
            var right = Assert.Single(saved.Slides[2].TextBoxes, box => box.Text == "Track the next action to closure.");
            results[item.Item1] = (body.Paragraphs[0].Runs[0].FontSize!.Value.ToPoints(),
                body.Width.ToPoints(), picture.Width.ToPoints(), left.Width.ToPoints(), right.Width.ToPoints());
            if (item.Item1 == "Purpose")
                Assert.Contains(saved.Slides[1].TextBoxes, box => string.Equals(box.Text, "Technical proposal", StringComparison.OrdinalIgnoreCase));
        }
        Assert.True(results["Compact"].Size < results["Default"].Size);
        Assert.True(results["Compact"].TextWidth > results["Default"].TextWidth);
        Assert.True(results["Visual"].ImageWidth > results["Default"].ImageWidth);
        Assert.True(results["Style"].Size < results["Default"].Size);
        Assert.Equal(results["Default"], results["ExplicitDefault"]);
        Assert.Equal(results["Design"], results["Repeat"]);
        Assert.NotEqual(results["Default"].Left, results["Design"].Left);
    }
}
