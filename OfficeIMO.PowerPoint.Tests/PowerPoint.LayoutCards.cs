using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointLayoutCardTests {
    private static readonly string[] Titles = { "Discover", "Stabilize", "Measure", "Deliver", "Review", "Improve" };

    [Fact]
    public void SavedCardHeadingsShareABaselineDespiteDifferentBodyLengths() {
        using var document = Compose(4, longBody: false);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        var snapshot = Assert.Single(saved.Slides).CreateVisualSnapshot();
        var spans = snapshot.TextSpans;
        double[] headings = Titles.Take(4).Select(title => Assert.Single(spans, span => span.Text == title).Y).ToArray();

        Assert.InRange(headings.Max() - headings.Min(), 0D, 0.1D);
        Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "powerpoint-text-overflow" &&
            diagnostic.Source?.IndexOf("Card Text", StringComparison.Ordinal) >= 0);
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void SavedLongCardTextStaysBelowItsNumberOnEveryContinuation() {
        using var document = Compose(6, longBody: true);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        int found = 0;
        foreach (PowerPointSlide slide in saved.Slides) {
            var spans = slide.CreateVisualSnapshot().TextSpans;
            for (int index = 0; index < Titles.Length; index++) {
                var heading = spans.FirstOrDefault(span => span.Text == Titles[index]);
                if (heading == null) continue;
                var number = Assert.Single(spans, span => span.Text == (index + 1).ToString("00"));
                PowerPointTextBox frame = Assert.Single(slide.TextBoxes,
                    box => box.Text.StartsWith(Titles[index], StringComparison.Ordinal));
                PowerPointTextBox numberFrame = Assert.Single(slide.TextBoxes,
                    box => box.Text == (index + 1).ToString("00"));
                Assert.True(frame.TopPoints >= numberFrame.TopPoints + numberFrame.HeightPoints,
                    Titles[index] + " has a text frame above or over its card number.");
                // Snapshot baselines use PDF coordinates, increasing upwards.
                Assert.True(heading.Y + heading.FontSize <= number.Y,
                    Titles[index] + " starts above or overlaps its card number.");
                Assert.Contains(spans, span => span.Text.IndexOf("Complete" + index, StringComparison.Ordinal) >= 0);
                found++;
            }
        }
        Assert.Equal(6, found);
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void OversizedCardBodyContinuesWithoutLosingLinesOrCrossingItsPanel() {
        using var document = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern });
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#205040", "oversized-card"));
        options.ApplyTheme = false;
        string[] lines = Enumerable.Range(1, 20).Select(i => "Evidence line " + i.ToString("00")).ToArray();
        document.Compose(new PowerPointDeckPlan().AddContent("Evidence", new[] {
            new PowerPointLayoutCard("Review", string.Join("\n", lines))
        }, PowerPointLayout.Cards), options);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.True(saved.Slides.Count > 1, "A single oversized card must continue before it leaves its panel.");
        foreach (string line in lines) Assert.Equal(1, saved.Slides.SelectMany(s => s.TextBoxes).Count(b => b.Text.Contains(line, StringComparison.Ordinal)));
        foreach (PowerPointSlide slide in saved.Slides) {
            var frame = Assert.Single(slide.TextBoxes, box => box.Text.StartsWith("Review", StringComparison.Ordinal));
            var panel = Assert.Single(slide.Shapes, shape => shape.Name != null && System.Text.RegularExpressions.Regex.IsMatch(shape.Name, @"^Card \d+$"));
            Assert.True(frame.TopPoints + frame.HeightPoints <= panel.TopPoints + panel.HeightPoints + 0.1D);
            Assert.DoesNotContain(slide.CreateVisualSnapshot().Diagnostics, d => d.Code == "powerpoint-text-overflow" && d.Source?.Contains("Card Text", StringComparison.Ordinal) == true);
        }
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void ALongCardHeadingReducesDensityBeforeShrinkingTheRowBelowItsMinimum() {
        using var document = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern });
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#205040", "long-heading"));
        options.ApplyTheme = false;
        string[] titles = { new string('W', 15), "Keep", "Read", "Act" };
        document.Compose(new PowerPointDeckPlan().AddContent("Responsibilities", titles.Select(title => new PowerPointLayoutCard(title, "One short line.")), PowerPointLayout.Cards), options);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        foreach (string title in titles) {
            var span = Assert.Single(saved.Slides.SelectMany(slide => slide.CreateVisualSnapshot().TextSpans), item => item.Text == title);
            Assert.True(span.FontSize >= 16D, title + " fell below the heading minimum: " + span.FontSize);
        }
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CardContinuationsPreserveParagraphsBulletsAndLogicalNumbering(bool bullets) {
        using var document = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Dark });
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#205040", "card-content"));
        options.ApplyTheme = false;
        string[] entries = Enumerable.Range(1, 45).Select(i => "Preserve the current delivery evidence Marker" + i.ToString("00") + ".").ToArray();
        document.Compose(new PowerPointDeckPlan().AddContent("Decisions", new[] {
            new PowerPointLayoutCard("Review", bullets ? null : string.Join(" ", entries), bullets ? entries : null),
            new PowerPointLayoutCard("Next", "Keep the next logical card.")
        }, PowerPointLayout.Cards), options);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.True(saved.Slides.Count > 2);
        foreach (int index in Enumerable.Range(1, 45)) {
            string marker = "Marker" + index.ToString("00") + ".";
            Assert.Equal(1, saved.Slides.SelectMany(slide => slide.TextBoxes).Count(box => box.Text.Contains(marker, StringComparison.Ordinal)));
        }
        foreach (PowerPointSlide slide in saved.Slides) {
            var frame = Assert.Single(slide.TextBoxes, box => box.Name?.StartsWith("Card Text", StringComparison.Ordinal) == true);
            var number = Assert.Single(slide.TextBoxes, box => box.Name?.StartsWith("Card Number", StringComparison.Ordinal) == true);
            Assert.Equal(frame.Text.StartsWith("Review", StringComparison.Ordinal) ? "01" : "02", number.Text);
            var panel = Assert.Single(slide.Shapes, shape => shape.Name != null && System.Text.RegularExpressions.Regex.IsMatch(shape.Name, @"^Card \d+$"));
            Assert.True(frame.TopPoints + frame.HeightPoints <= panel.TopPoints + panel.HeightPoints + 0.1D);
            Assert.DoesNotContain(slide.CreateVisualSnapshot().Diagnostics, d => d.Code == "powerpoint-text-overflow" && d.Source?.Contains("Card Text", StringComparison.Ordinal) == true);
        }
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void AnUnbreakableHeadingThatCannotFitAtMinimumSizeFailsWithoutAddingSlides() {
        using var document = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern });
        var original = document.AddSlide();
        Assert.Throws<ArgumentException>(() => document.Compose(new PowerPointDeckPlan().AddContent("Decisions", new[] {
            new PowerPointLayoutCard("First", "Keep this evidence."), new PowerPointLayoutCard(new string('W', 200), "Too wide.")
        }, PowerPointLayout.Cards), PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#205040", "invalid-card"))));
        Assert.Same(original, Assert.Single(document.Slides));
    }

    private static PowerPointPresentation Compose(int count, bool longBody) {
        var document = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern });
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#205040", "card-layout"));
        options.ApplyTheme = false;
        var cards = Titles.Take(count).Select((title, index) => new PowerPointLayoutCard(title,
            string.Join(" ", Enumerable.Repeat("Keep reliable evidence for every source and follow the agreed delivery plan.",
                longBody ? 8 : index + 1)) + " Complete" + index)).ToArray();
        document.Compose(new PowerPointDeckPlan().AddContent("Planning options", cards, PowerPointLayout.Cards), options);
        return document;
    }
}
