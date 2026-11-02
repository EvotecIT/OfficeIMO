using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointLayoutMetricTests {
    [Theory]
    [InlineData(PowerPointLayout.Kpi)]
    [InlineData(PowerPointLayout.CaseStudy)]
    public void LowerIsBetterUsesTheSameChangeInMetricAndCaseStudyLayouts(PowerPointLayout layout) {
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern });
        var metric = PowerPointLayoutKpi.FromDelta("182 ms", "Median response", -0.14D, "%", PowerPointKpiDirection.Down);
        PowerPointLayoutBlock[] blocks = layout == PowerPointLayout.CaseStudy
            ? new PowerPointLayoutBlock[] { new PowerPointLayoutCard("Routing", "Reduce processing time."), metric }
            : new PowerPointLayoutBlock[] { metric };
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2563EB", "metric-change"));
        options.ApplyTheme = false;
        document.Compose(new PowerPointDeckPlan().AddContent("Service health", blocks, layout), options);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        var change = Assert.Single(Assert.Single(saved.Slides).TextBoxes, box => box.Text == "▼ -14%");
        Assert.Equal("▼ -14%", change.Text);
        Assert.Equal(OfficeTheme.Modern.Palette.Success.ToRgbHex(),
            Assert.Single(change.Paragraphs.SelectMany(paragraph => paragraph.Runs)).Color, ignoreCase: true);
        Assert.DoesNotContain(Assert.Single(saved.Slides).CreateVisualSnapshot().Diagnostics,
            diagnostic => diagnostic.Code == "powerpoint-text-overflow");
        Assert.Empty(saved.ValidateDocument());
    }
}
