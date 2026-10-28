using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointCardAdmissionTests {
    [Theory]
    [InlineData("Body")]
    [InlineData("Heading")]
    [InlineData("Count")]
    public void CustomCardRegionsRejectImpossibleContentWithoutLeavingPartialSlides(string boundary) {
        using var document = PowerPointPresentation.Create();
        PowerPointSlide original = document.AddSlide();
        var cards = boundary == "Count" ? Enumerable.Range(1, 7).Select(i => new PowerPointCardContent("Card " + i)).ToArray()
            : new[] { new PowerPointCardContent(boundary == "Heading" ? new string('W', 200) : "Review") {
                Text = boundary == "Body" ? string.Join("\n", Enumerable.Range(1, 20).Select(i => "Evidence line " + i)) : "One line."
            } };
        var plan = new PowerPointDeckPlan().AddCustom("Review", context => context.AddCardGrid(cards));
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#205040", "custom-admission"));
        options.ApplyTheme = false;
        Assert.Throws<ArgumentException>(() => document.Compose(plan, options));
        Assert.Same(original, Assert.Single(document.Slides));
    }

    [Fact]
    public void TemplateCardContinuationsFitBelowTheActualTitlePlaceholder() {
        using var template = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern });
        foreach (var layout in template.GetSlideLayouts()) {
            if (template.GetLayoutPlaceholderTextBox(layout.MasterIndex, layout.LayoutIndex, PowerPointPlaceholderType.Title) != null)
                template.SetLayoutPlaceholderBounds(layout.MasterIndex, layout.LayoutIndex, PowerPointPlaceholderType.Title,
                    PowerPointLayoutBox.FromPoints(48, 50, 864, 150));
        }
        using var document = PowerPointPresentation.Load(new MemoryStream(template.ToBytes()));
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#205040", "template-admission"));
        options.ApplyTheme = false;
        string[] lines = Enumerable.Range(1, 20).Select(i => "Evidence line " + i.ToString("00")).ToArray();
        document.Compose(new PowerPointDeckPlan().AddContent("Evidence", new[] {
            new PowerPointLayoutCard("Review", string.Join("\n", lines))
        }, PowerPointLayout.Cards), options);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.True(saved.Slides.Count > 1);
        foreach (PowerPointSlide slide in saved.Slides) {
            var frame = Assert.Single(slide.TextBoxes, box => box.Name?.StartsWith("Card Text", StringComparison.Ordinal) == true);
            Assert.True(frame.TopPoints >= 200D);
            double declaredHeight = frame.Paragraphs.Sum(paragraph =>
                (paragraph.SpaceBeforePoints ?? 0D) + (paragraph.SpaceAfterPoints ?? 0D) +
                (paragraph.InlineNodes.Count(node => node.Kind == PowerPointParagraphInlineKind.LineBreak) + 1) * (paragraph.LineSpacingPoints ?? 0D));
            Assert.True(declaredHeight <= frame.HeightPoints + 0.1D,
                "The saved card's declared line spacing needs " + declaredHeight + "pt in a " + frame.HeightPoints + "pt frame.");
            Assert.DoesNotContain(slide.CreateVisualSnapshot().Diagnostics,
                d => d.Code == "powerpoint-text-overflow" && d.Source?.Contains("Card Text", StringComparison.Ordinal) == true);
        }
        foreach (string line in lines) Assert.Equal(1, saved.Slides.SelectMany(slide => slide.TextBoxes).Count(box => box.Text.Contains(line, StringComparison.Ordinal)));
        using var pdf = saved.ToPdf(null, out var report);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "text-box-overflow");
        Assert.Empty(saved.ValidateDocument());
    }
}
