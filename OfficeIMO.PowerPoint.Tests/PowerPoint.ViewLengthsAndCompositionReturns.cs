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
            var title = context.AddTitle("Original", "Subtitle");
            Assert.Equal(2, title.Count);
            Assert.All(title, shape => Assert.Contains(shape, context.Slide.Shapes));
            Assert.IsType<PowerPointTextBox>(title[0]).Text = "Edited title";
            var callout = context.AddCalloutBand("Original callout", context.ContentArea());
            Assert.Equal(3, callout.Count);
            Assert.All(callout, shape => Assert.DoesNotContain(shape, title));
            Assert.IsType<PowerPointTextBox>(callout[2]).Text = "Edited callout";
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
    }
}
