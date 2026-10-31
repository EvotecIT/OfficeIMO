using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public class PowerPointTextColorOpacityTests {
    [Theory]
    [InlineData("run")]
    [InlineData("box")]
    [InlineData("highlight")]
    public void TranslucentTextColorsUseRgbAndNativeAlphaAfterSave(string target) {
        using var document = PowerPointDocument.Create();
        var box = document.AddSlide().AddTextBox("Translucent text");
        var run = box.Paragraphs.Single().Runs.Single();
        if (target == "box") box.Color = "33669980";
        else if (target == "run") run.Color = "33669980";
        else run.HighlightColor = "33669980";

        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        var actual = saved.Slides.Single().TextBoxes.Single().Paragraphs.Single().Runs.Single();
        var properties = actual.Run.RunProperties!;
        A.RgbColorModelHex rgb = target == "highlight"
            ? properties.GetFirstChild<A.Highlight>()!.GetFirstChild<A.RgbColorModelHex>()!
            : properties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!;
        Assert.Equal("336699", rgb.Val!.Value);
        Assert.Equal(50196, Assert.Single(rgb.Elements<A.Alpha>()).Val!.Value);
    }
}
