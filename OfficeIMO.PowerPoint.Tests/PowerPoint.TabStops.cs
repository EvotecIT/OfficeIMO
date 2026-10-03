using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public class PowerPointTabStopsTests {
    [Fact]
    public void ParagraphTabStopsReplaceAtomicallyAndSurviveSave() {
        using var presentation = PowerPointPresentation.Create();
        var paragraph = presentation.AddSlide().AddTextBox("A\tB").Paragraphs.Single();
        paragraph.SetTabStops(new[] { new PowerPointTabStop(18), new PowerPointTabStop(72, PowerPointTabAlignment.Decimal) });
        Assert.Throws<ArgumentException>(() => paragraph.SetTabStops(new PowerPointTabStop[] { new(36), null! }));
        Assert.Equal(new double[] { 18, 72 }, paragraph.TabStops.Select(t => t.PositionPoints));
        Assert.Empty(presentation.ValidateDocument());
        using var output = new MemoryStream();
        presentation.Save(output); output.Position = 0;
        using var reopened = PowerPointPresentation.Load(output);
        var restored = reopened.Slides.Single().TextBoxes.Single().Paragraphs.Single();
        Assert.Equal(PowerPointTabAlignment.Decimal, restored.TabStops[1].Alignment);
        restored.SetTabStops(Array.Empty<PowerPointTabStop>());
        Assert.Empty(restored.TabStops);
        Assert.Equal("A\tB", restored.Text);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PowerPointTabStop(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PowerPointTabStop(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PowerPointTabStop(18, (PowerPointTabAlignment)9));
    }
}
