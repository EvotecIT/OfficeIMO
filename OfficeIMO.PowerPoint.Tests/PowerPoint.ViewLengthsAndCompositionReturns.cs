using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointViewLengthsAndCompositionReturns {
    [Fact]
    public void ViewLengthsAndReturnedGuidesRoundTripThroughTheNativePackage() {
        using var document = PowerPointDocument.Create();
        document.AddSlide();
        document.SetGridSpacing(OfficeLength.Emus(12345), OfficeLength.Points(3));
        var guide = document.AddGuide(PowerPointGuideOrientation.Vertical, OfficeLength.Emus(-12345));
        Assert.Equal(OfficeLength.Emus(-12345), guide.Position);
        var columns = document.AddColumnGuides(2, OfficeLength.Points(10), OfficeLength.Points(5));
        Assert.Equal(4, columns.Count);
        Assert.All(columns, item => Assert.Equal(PowerPointGuideOrientation.Vertical, item.Orientation));
        var rows = document.AddRowGuides(2, OfficeLength.Points(10), OfficeLength.Points(5), includeOuterEdges: false);
        Assert.Equal(2, rows.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetGridSpacing(OfficeLength.Points(7), OfficeLength.Emus(long.MaxValue)));
        Assert.Equal(OfficeLength.Emus(12345), document.GridSpacingX);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using (var package = PresentationDocument.Open(output, false)) {
            var view = package.PresentationPart!.ViewPropertiesPart!.ViewProperties!;
            Assert.Equal(12345, view.GridSpacing!.Cx!.Value);
            Assert.Equal(38100, view.GridSpacing.Cy!.Value);
        }
        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        Assert.Equal(document.GridSpacingY, loaded.GridSpacingY);
        Assert.Equal(guide.Position, loaded.GetGuides()[0].Position);
        Assert.Equal(1 + columns.Count + rows.Count, loaded.GetGuides().Count);
    }

    [Fact]
    public void CompositeAddsReturnOnlyTheirNewAttachedEditableShapes() {
        using var document = PowerPointDocument.Create();
        var plan = new PowerPointDeckPlan().AddCustom("Editable composition", context => {
            var beforeTitle = context.Slide.Shapes.ToArray();
            var title = context.AddTitle("Original", "Subtitle");
            Assert.Equal(context.Slide.Shapes.Except(beforeTitle), title);
            Assert.All(title, shape => Assert.Contains(shape, context.Slide.Shapes));
            Assert.Single(title.OfType<PowerPointTextBox>(), box => box.Text == "Original").Text = "Edited title";
            var beforeCallout = context.Slide.Shapes.ToArray();
            var callout = context.AddCalloutBand("Original callout", context.ContentArea());
            Assert.Equal(context.Slide.Shapes.Except(beforeCallout), callout);
            Assert.All(callout, shape => Assert.DoesNotContain(shape, title));
            Assert.Single(callout.OfType<PowerPointTextBox>(), box => box.Text == "Original callout").Text = "Edited callout";
            var explicitText = context.AddTextBlock("Physical font", context.ContentArea(), OfficeLength.Inches(0.1875));
            Assert.Equal(OfficeLength.Points(13.5), explicitText.FontSize);
            var defaultText = context.AddTextBlock("Default font", context.ContentArea());
            Assert.Equal(OfficeLength.Points(12), defaultText.FontSize);
            int count = context.Slide.Shapes.Count;
            Assert.Throws<ArgumentOutOfRangeException>(() => context.AddTextBlock("Invalid font", context.ContentArea(), OfficeLength.Zero));
            Assert.Equal(count, context.Slide.Shapes.Count);
            Assert.Empty(context.AddMetricStrip(Array.Empty<PowerPointMetric>(), context.ContentArea()));
        });
        var options = PowerPointCompositionOptions.FromBrief(PowerPointDesignBrief.FromBrand("#008C95", "composition-return-proof"));
        options.SelectBestAlternative = false;
        document.Compose(plan, options);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        var text = loaded.Slides.Single().TextBoxes.Select(box => box.Text).ToArray();
        Assert.Contains("Edited title", text);
        Assert.Contains("Edited callout", text);
        var textBoxes = loaded.Slides.Single().TextBoxes;
        Assert.Equal(OfficeLength.Points(13.5), textBoxes.Single(box => box.Text == "Physical font").FontSize);
        Assert.Equal(OfficeLength.Points(12), textBoxes.Single(box => box.Text == "Default font").FontSize);
    }
}
