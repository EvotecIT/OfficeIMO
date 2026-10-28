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
