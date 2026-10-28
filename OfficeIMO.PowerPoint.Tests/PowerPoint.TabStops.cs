using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public class PowerPointTabStopsTests {
    [Fact]
    public void ParagraphTabStopsReplaceAtomicallyAndSurviveSave() {
        using var presentation = PowerPointDocument.Create();
        var paragraph = presentation.AddSlide().AddTextBox("A\tB").Paragraphs.Single();
        paragraph.SetTabStops(new[] { new PowerPointTabStop(global::OfficeIMO.OfficeLength.Points(18)), new PowerPointTabStop(global::OfficeIMO.OfficeLength.Points(72), PowerPointTabAlignment.Decimal) });
        Assert.Throws<ArgumentException>(() => paragraph.SetTabStops(new PowerPointTabStop[] { new(global::OfficeIMO.OfficeLength.Points(36)), null! }));
        Assert.Equal(new double[] { 18, 72 }, paragraph.TabStops.Select(t => t.Position.ToPoints()));
        Assert.Empty(presentation.ValidateDocument());
        using var output = new MemoryStream();
        presentation.Save(output); output.Position = 0;
        using var reopened = PowerPointDocument.Load(output);
        var restored = reopened.Slides.Single().TextBoxes.Single().Paragraphs.Single();
        Assert.Equal(PowerPointTabAlignment.Decimal, restored.TabStops[1].Alignment);
        Assert.Equal(global::OfficeIMO.OfficeLength.Inches(1), restored.TabStops[1].Position);
        restored.SetTabStops(Array.Empty<PowerPointTabStop>());
        Assert.Empty(restored.TabStops);
        Assert.Equal("A\tB", restored.Text);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PowerPointTabStop(global::OfficeIMO.OfficeLength.Emus((long)int.MaxValue + 1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PowerPointTabStop(global::OfficeIMO.OfficeLength.Points(-1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PowerPointTabStop(global::OfficeIMO.OfficeLength.Points(18), (PowerPointTabAlignment)9));
    }
}
