using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointLayoutTypographyTests {
    [Fact]
    public void SavedDisplayTitleRetainsItsEmBasedLeadingInTheVisualSnapshot() {
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Boardroom });
        var plan = new PowerPointDeckPlan().AddContent("FIRST\nSECOND", Array.Empty<PowerPointLayoutBlock>());
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#205040", "display-leading"));
        options.ApplyTheme = false;
        document.Compose(plan, options);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        var spans = saved.Slides[0].CreateVisualSnapshot().TextSpans.Where(span =>
            span.Text.Contains("FIRST") || span.Text.Contains("SECOND")).ToArray();
        double[] lines = spans.Select(span => Math.Round(span.Y, 3)).Distinct().OrderBy(value => value).ToArray();
        Assert.Equal(2, lines.Length);
        double fontSize = spans.Max(span => span.FontSize);
        // Display leading is one em; writing and reading the file must not compress it to five sixths of an em.
        Assert.InRange(lines[1] - lines[0], fontSize - 0.1D, fontSize + 0.1D);
        Assert.Empty(saved.ValidateDocument());
    }
}
