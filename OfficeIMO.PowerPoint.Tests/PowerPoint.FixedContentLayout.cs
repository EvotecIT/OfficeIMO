using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointFixedContentLayoutTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FixedFiguresUseAContentHeadingUnlessAnOpeningTitleIsExplicitlyRequested(bool openingTitle) {
        using var document = PowerPointDocument.Create();
        using var materialization = new PowerPointLayoutMaterialization(document);
        materialization.Slide.AddShape(OfficeIMO.Drawing.OfficePresetShapeType.Rectangle,
            OfficeLength.Points(60), OfficeLength.Points(140), OfficeLength.Points(420), OfficeLength.Points(200), "Fixed content");
        var plan = new PowerPointContentPlanSlide("Manual composition", Array.Empty<PowerPointLayoutBlock>(),
            openingTitle ? PowerPointLayout.Title : PowerPointLayout.Auto);
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#008C95", "fixed-content"));
        materialization.Complete(plan, options);
        using var bytes = new MemoryStream();
        document.Save(bytes);
        bytes.Position = 0;
        using var saved = PowerPointDocument.Load(bytes);
        Assert.Empty(saved.ValidateDocument());
        var slide = Assert.Single(saved.Slides);
        var fixedContent = Assert.Single(slide.Shapes, shape => shape.Name == "Fixed content");
        var title = Assert.Single(slide.TextBoxes, box => box.Text == "Manual composition");
        Assert.Equal(762000L, fixedContent.Left.ToEmus());
        Assert.Equal(1778000L, fixedContent.Top.ToEmus());
        Assert.Equal(5334000L, fixedContent.Width.ToEmus());
        Assert.Equal(2540000L, fixedContent.Height.ToEmus());
        if (openingTitle) Assert.True(title.Top > fixedContent.Top);
        else Assert.True(title.Top + title.Height <= fixedContent.Top,
            "Automatic layout must keep the heading above deliberately placed content.");
    }
}
