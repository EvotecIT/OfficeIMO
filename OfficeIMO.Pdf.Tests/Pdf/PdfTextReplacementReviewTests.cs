using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextReplacementReviewTests {
    [Fact]
    public void ClickingTextSelectsItsWordWithoutAdjacentWhitespaceOrWords() {
        PdfDocument source = Create("Before Account After");
        PdfTextMatch match = Assert.Single(source.Text.Find("Account"));
        PdfPageInteractionMap map = source.Render.Interactions(1);
        var word = map.SelectWord((match.VisualBounds.TopLeft.X + match.VisualBounds.BottomRight.X) / 2,
            (match.VisualBounds.TopLeft.Y + match.VisualBounds.BottomRight.Y) / 2);
        Assert.Equal("Account", string.Concat(word.Select(region => region.Text)));
        Assert.Empty(map.SelectWord(1, 1));
    }

    [Fact]
    public void SelectedOccurrencesShareOneRewriteAndPreserveUnselectedText() {
        PdfDocument source = Create("Account alpha Account beta Account gamma");
        PdfTextEditResult result = source.Text.ReplaceSelected("Account", "Record", new[] { 0, 2 });
        Assert.Equal(2, result.AffectedCount);
        Assert.Equal(2, result.Document.Text.Find("Record").Count);
        Assert.Single(result.Document.Text.Find("Account"));
        Assert.Contains("beta", result.Document.Read().Text);
        Assert.Equal(3, source.Text.Find("Account").Count);
        Assert.Throws<ArgumentException>(() => source.Text.ReplaceSelected("Account", "Record", new[] { 0, 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => source.Text.ReplaceSelected("Account", "Record", new[] { 3 }));
    }

    [Fact]
    public void MatchFitPolicyRejectsOverflowOrShrinksOnlyReplacement() {
        PdfDocument source = Create("Before Account After");
        PdfTextMatch match = Assert.Single(source.Text.Find("Account"));
        const string replacement = "Longer account";
        Assert.Throws<NotSupportedException>(() => source.Text.Replace(match, replacement,
            new PdfTextEditOptions { RegionWidthPolicy = PdfTextRegionWidthPolicy.RejectOverflow }));
        PdfTextEditResult result = source.Text.Replace(match, replacement, new PdfTextEditOptions {
            RegionWidthPolicy = PdfTextRegionWidthPolicy.ShrinkToFit, MinimumFontSize = 1
        });
        PdfTextMatch changed = Assert.Single(result.Document.Text.Find(replacement));
        Assert.True(changed.FontSize < match.FontSize);
        Assert.True(changed.Width <= match.Width + 0.1);
        Assert.Equal(Assert.Single(source.Text.Find("Before")).FontSize,
            Assert.Single(result.Document.Text.Find("Before")).FontSize, precision: 3);
        Assert.Equal(Assert.Single(source.Text.Find("After")).FontSize,
            Assert.Single(result.Document.Text.Find("After")).FontSize, precision: 3);
        Assert.NotEmpty(result.Warnings);
    }

    private static PdfDocument Create(string text) => PdfDocument.Create(compose =>
        compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(text))))));
}
